"""
LÆS FILINDHOLD

PDF-filer:
- læses altid først med pypdf
- alle sider læses
- Azure bruges kun, hvis PDF-filen mangler tekstlag

Billeder:
- kræver Azure Vision

Vedhæftede mails:
- hentes som bytes
- indholdet læses ikke

Filtekst og bytes gemmes kun i hukommelsen.
"""

from io import BytesIO
from pathlib import Path

from pypdf import PdfReader

from q_haderslev_vbo.automation_server.ats_update_item_data import (
    update_item_data,
)

from kald_azure_vision import (
    azure_vision_is_configured,
    read_file_with_azure_vision,
)


# -------------------------------------------------
# KONSTANTER
# -------------------------------------------------

MAX_AZURE_ATTEMPTS_PER_ITEM = 50

STATE_AZURE_LIMIT_REACHED = (
    "3.2 - Maksimalt antal Azure-forsøg nået"
)

IMAGE_SUFFIXES = {
    ".bmp",
    ".gif",
    ".jpeg",
    ".jpg",
    ".png",
    ".tif",
    ".tiff",
    ".webp",
}


# -------------------------------------------------
# LÆS PDF MED PYPDF
# -------------------------------------------------

def read_pdf_with_pypdf(file_bytes):
    """
    Læser alle PDF-sider direkte fra bytes.
    """

    reader = PdfReader(
        BytesIO(file_bytes)
    )

    page_texts = []

    for page in reader.pages:
        page_text = (
            page.extract_text()
            or ""
        )

        page_texts.append(
            page_text
        )

    return "\n".join(
        page_texts
    ).strip()


# -------------------------------------------------
# HAR PDF TEKST?
# -------------------------------------------------

def pdf_has_text(text):
    """
    Kontrollerer om pypdf fandt bogstaver eller tal.
    """

    return any(
        character.isalnum()
        for character in str(text or "")
    )


# -------------------------------------------------
# AZURE COUNTER
# -------------------------------------------------

def get_azure_attempt_count(item):
    """
    Henter Azure-counteren fra item.data["box"].
    """

    return int(
        (
            item.data
            .get("box", {})
            .get("azure_attempt_count", 0)
        )
        or 0
    )


def increment_azure_attempt_count(item):
    """
    Forøger counteren før Azure-kaldet.
    """

    new_count = (
        get_azure_attempt_count(item)
        + 1
    )

    update_item_data(
        item.data,
        box_updates={
            "azure_attempt_count": new_count,
        },
        item=item,
    )

    return new_count


# -------------------------------------------------
# LÆS ÉN VEDHÆFTNING
# -------------------------------------------------

def read_one_attachment(
    attachment,
    item,
):
    """
    Læser én vedhæftning til hukommelsen.
    """

    filename = (
        attachment.get("name")
        or "Vedhæftet fil"
    )

    file_bytes = attachment.get(
        "content_bytes"
    )

    attachment_type = attachment.get(
        "attachment_type"
    )

    item_type = attachment.get(
        "item_type"
    )

    suffix = (
        Path(filename)
        .suffix
        .casefold()
    )

    result = {
        "name": filename,
        "text": "",
        "read_method": "skipped",
        "azure_used": False,
        "success": True,
        "error": None,
    }

    if not isinstance(file_bytes, bytes):
        result["success"] = False
        result["error"] = (
            "Vedhæftningen mangler content_bytes"
        )

        return result

    # Vedhæftede mails skal ikke åbnes.
    if (
        attachment_type == "itemAttachment"
        or item_type == "message"
        or suffix == ".eml"
    ):
        result["read_method"] = (
            "attached_email_skipped"
        )

        return result

    local_text = ""
    needs_azure = False

    # -------------------------------------------------
    # PDF
    # -------------------------------------------------

    if suffix == ".pdf":
        try:
            local_text = read_pdf_with_pypdf(
                file_bytes
            )

            if pdf_has_text(local_text):
                result["text"] = local_text
                result["read_method"] = "pypdf"

                return result

            needs_azure = True

        except Exception as error:
            needs_azure = True
            result["error"] = str(error)

    # -------------------------------------------------
    # BILLEDE
    # -------------------------------------------------

    elif suffix in IMAGE_SUFFIXES:
        needs_azure = True

    # -------------------------------------------------
    # IKKE UNDERSTØTTET
    # -------------------------------------------------

    else:
        result["read_method"] = (
            "unsupported_file_type"
        )

        return result

    if not needs_azure:
        return result

    # Azure er endnu ikke koblet på.
    if not azure_vision_is_configured():
        result["text"] = local_text
        result["read_method"] = (
            "azure_not_configured"
        )

        return result

    current_count = get_azure_attempt_count(
        item
    )

    if (
        current_count
        >= MAX_AZURE_ATTEMPTS_PER_ITEM
    ):
        update_item_data(
            item.data,
            state=STATE_AZURE_LIMIT_REACHED,
            item=item,
        )

        result["text"] = local_text
        result["read_method"] = (
            "azure_limit_reached"
        )

        return result

    increment_azure_attempt_count(item)

    try:
        azure_text = (
            read_file_with_azure_vision(
                filename=filename,
                file_bytes=file_bytes,
            )
        )

        result["text"] = azure_text or ""
        result["read_method"] = "azure_vision"
        result["azure_used"] = True
        result["error"] = None

    except Exception as error:
        result["success"] = False
        result["read_method"] = (
            "azure_vision_failed"
        )
        result["azure_used"] = True
        result["error"] = str(error)

    return result


# -------------------------------------------------
# LÆS ALLE VEDHÆFTNINGER
# -------------------------------------------------

def read_attachment_contents(
    attachments,
    item,
):
    """
    Læser alle vedhæftninger.

    En fejl på én fil stopper ikke de øvrige filer.
    """

    results = []

    for attachment in attachments:
        try:
            result = read_one_attachment(
                attachment=attachment,
                item=item,
            )

        except Exception as error:
            result = {
                "name": (
                    attachment.get("name")
                    or "Vedhæftet fil"
                ),
                "text": "",
                "read_method": "failed",
                "azure_used": False,
                "success": False,
                "error": str(error),
            }

        results.append(result)

    return results