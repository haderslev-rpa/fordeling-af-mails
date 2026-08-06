"""
LÆS FILINDHOLD

PDF-filer:
- læses kun med pypdf
- alle sider læses
- sendes aldrig til Azure
- hvis PDF-filen mangler tekstlag, returneres tom tekst

Billedfiler:
- sendes til Azure Vision

Vedhæftede mails:
- læses ikke

Andre filtyper:
- springes over

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

from proces_konfiguration import (
    MAX_AZURE_ATTEMPTS_PER_ITEM,
)


# -------------------------------------------------
# BILLEDTYPER
# -------------------------------------------------

IMAGE_SUFFIXES = {
    ".bmp",
    ".jpeg",
    ".jpg",
    ".png",
    ".tif",
    ".tiff",
}


# -------------------------------------------------
# LÆS PDF MED PYPDF
# -------------------------------------------------

def read_pdf_with_pypdf(
    file_bytes,
):
    """
    Læser tekst fra alle PDF-sider.

    PDF-filen læses direkte fra bytes.
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
    Kontrollerer om PDF-teksten indeholder
    bogstaver eller tal.
    """

    return any(
        character.isalnum()
        for character in str(
            text or ""
        )
    )


# -------------------------------------------------
# HENT AZURE COUNTER
# -------------------------------------------------

def get_azure_attempt_count(
    item,
):
    """
    Henter Azure-counteren fra box.
    """

    return int(
        (
            item.data
            .get("box", {})
            .get(
                "azure_attempt_count",
                0,
            )
        )
        or 0
    )


# -------------------------------------------------
# FORØG AZURE COUNTER
# -------------------------------------------------

def increment_azure_attempt_count(
    item,
):
    """
    Forøger Azure-counteren før Azure-kaldet.
    """

    new_count = (
        get_azure_attempt_count(
            item
        )
        + 1
    )

    item.data["box"][
        "azure_attempt_count"
    ] = new_count

    update_item_data(
        item.data,
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
    Læser én vedhæftning.

    Returnerer udlæst tekst og metadata
    i hukommelsen.
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
        "azure_limit_reached": False,
        "success": True,
        "error": None,
    }

    if not isinstance(
        file_bytes,
        bytes,
    ):
        result["success"] = False
        result["error"] = (
            "Vedhæftningen mangler content_bytes"
        )

        return result

    # -------------------------------------------------
    # VEDHÆFTET MAIL
    # -------------------------------------------------

    if (
        attachment_type == "itemAttachment"
        or item_type == "message"
        or suffix == ".eml"
    ):
        result["read_method"] = (
            "attached_email_skipped"
        )

        return result

    # -------------------------------------------------
    # PDF LÆSES KUN MED PYPDF
    # -------------------------------------------------

    if suffix == ".pdf":
        try:
            pdf_text = read_pdf_with_pypdf(
                file_bytes
            )

            result["text"] = pdf_text
            result["read_method"] = "pypdf"

            if not pdf_has_text(
                pdf_text
            ):
                result["read_method"] = (
                    "pypdf_no_text_layer"
                )

            return result

        except Exception as error:
            result["success"] = False
            result["read_method"] = (
                "pypdf_failed"
            )
            result["error"] = str(
                error
            )

            return result

    # -------------------------------------------------
    # IKKE-BILLEDE SPRINGES OVER
    # -------------------------------------------------

    if suffix not in IMAGE_SUFFIXES:
        result["read_method"] = (
            "unsupported_file_type"
        )

        return result

    # -------------------------------------------------
    # AZURE ER IKKE KONFIGURERET
    # -------------------------------------------------

    if not azure_vision_is_configured():
        result["success"] = False
        result["read_method"] = (
            "azure_not_configured"
        )
        result["error"] = (
            "Azure Vision er ikke konfigureret"
        )

        return result

    # -------------------------------------------------
    # KONTROLLÉR AZURE-GRÆNSE
    # -------------------------------------------------

    current_count = (
        get_azure_attempt_count(
            item
        )
    )

    if (
        current_count
        >= MAX_AZURE_ATTEMPTS_PER_ITEM
    ):
        result["read_method"] = (
            "azure_limit_reached"
        )
        result["azure_limit_reached"] = True

        return result

    # Counteren gemmes før Azure-kaldet.
    increment_azure_attempt_count(
        item
    )

    # -------------------------------------------------
    # LÆS BILLEDE MED AZURE
    # -------------------------------------------------

    try:
        azure_text = (
            read_file_with_azure_vision(
                filename=filename,
                file_bytes=file_bytes,
            )
        )

        result["text"] = (
            azure_text
            or ""
        )
        result["read_method"] = (
            "azure_vision"
        )
        result["azure_used"] = True
        result["success"] = True
        result["error"] = None

    except Exception as error:
        result["success"] = False
        result["read_method"] = (
            "azure_vision_failed"
        )
        result["azure_used"] = True
        result["error"] = str(
            error
        )

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
                "azure_limit_reached": False,
                "success": False,
                "error": str(error),
            }

        results.append(
            result
        )

    return results