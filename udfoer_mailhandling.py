"""
UDFØR MAILHANDLING

Denne fil udfører den endelige Outlook-handling.

Logik:

Under 300 point:
- Overskriv alle kategorier.
- Videresend ikke.
- Flyt ikke.

Testregel:
- Overskriv alle kategorier.
- Videresend ikke.
- Flyt ikke.

Aktiv regel med mindst 300 point:
- Videresend, hvis destinationen er en mailadresse.
- Flyt, hvis destinationen er en Outlook-mappe.
- Udfør ingen handling, hvis destinationen er blank.

VIGTIGT:

- Kategorier overskrives med update_mail_categories.
- add_mail_category bruges ikke.
- PDF/Azure-logik findes ikke i denne fil.
- Funktionen returnerer resultatet til behandel.py.
- behandel.py sætter state efter en vellykket handling.
- main.py sætter status og status_code til sidst.
"""

from html import escape

from q_outlook_api.functionality.mail_api import (
    delete_mail,
    forward_mail,
    get_folders,
    move_mail,
    update_mail_categories,
)


# -------------------------------------------------
# ER DESTINATION EN MAILADRESSE?
# -------------------------------------------------

def destination_is_email(destination):
    """
    Kontrollerer om destinationen ligner
    en mailadresse.
    """

    destination_text = str(
        destination or ""
    ).strip()

    return bool(
        destination_text
        and "@" in destination_text
    )


# -------------------------------------------------
# FIND MAILMAPPE
# -------------------------------------------------

def find_folder_id(
    mailbox,
    folder_name,
):
    """
    Finder Graph-id'et for en Outlook-mappe.

    Destinationen kan være:

    Autosvar

    eller en fuld sti:

    Indbakke/Autosvar

    Hvis flere mapper har samme navn,
    skal den fulde sti bruges.
    """

    requested_folder = str(
        folder_name or ""
    ).strip()

    if not requested_folder:
        raise ValueError(
            "Navnet på destinationsmappen mangler."
        )

    requested_folder = (
        requested_folder
        .replace("\\", "/")
        .strip("/")
    )

    folders = get_folders(
        user_mail=mailbox,
        include_hidden=True,
        include_children=True,
    )

    requested_normalized = (
        requested_folder.casefold()
    )


    # -------------------------------------------------
    # 1. FIND PRÆCIST MATCH PÅ MAPPESTI
    # -------------------------------------------------

    path_matches = [
        folder
        for folder in folders
        if (
            str(
                folder.get("path")
                or ""
            )
            .strip("/")
            .casefold()
            == requested_normalized
        )
    ]

    if len(path_matches) == 1:
        return path_matches[0]["id"]

    if len(path_matches) > 1:
        raise ValueError(
            "Flere Outlook-mapper har "
            f"stien '{requested_folder}'."
        )


    # -------------------------------------------------
    # 2. FIND MATCH PÅ MAPPENAVN
    # -------------------------------------------------

    name_matches = [
        folder
        for folder in folders
        if (
            str(
                folder.get(
                    "display_name"
                )
                or folder.get(
                    "displayName"
                )
                or ""
            )
            .strip()
            .casefold()
            == requested_normalized
        )
    ]

    if len(name_matches) == 1:
        return name_matches[0]["id"]

    if len(name_matches) > 1:
        matching_paths = sorted(
            {
                folder.get("path")
                for folder in name_matches
                if folder.get("path")
            }
        )

        raise ValueError(
            f"Flere Outlook-mapper hedder "
            f"'{requested_folder}'. "
            "Brug den fulde mappesti i Excel. "
            f"Mulige stier: {matching_paths}"
        )


    # -------------------------------------------------
    # MAPPE IKKE FUNDET
    # -------------------------------------------------

    available_paths = sorted(
        {
            folder.get("path")
            for folder in folders
            if folder.get("path")
        }
    )

    raise ValueError(
        f"Outlook-mappen '{requested_folder}' "
        "blev ikke fundet. "
        f"Tilgængelige mapper: "
        f"{available_paths}"
    )

# -------------------------------------------------
# BYG KATEGORIER
# -------------------------------------------------

def build_rule_categories(
    rule_result,
    include_test=False,
):
    """
    Bygger den komplette kategoriliste.

    Listen overskriver alle eksisterende
    kategorier på mailen.

    Ved 0 point tilføjes regelnummeret ikke.
    """

    total_points = int(
        rule_result.get(
            "total_points",
            0,
        )
        or 0
    )

    rule_number = rule_result.get(
        "rule_number"
    )

    categories = []

    if include_test:
        categories.append(
            "Test"
        )

    categories.append(
        f"Point: {total_points}"
    )

    # Regelnummeret er kun relevant,
    # når reglen faktisk har givet point.
    if (
        total_points > 0
        and rule_number is not None
    ):
        categories.append(
            f"Nr: {rule_number}"
        )

    return categories

# -------------------------------------------------
# BYG REGELFORKLARING SOM HTML
# -------------------------------------------------

def build_rule_explanation(
    rule_result,
    mailbox,
):
    """
    Bygger den HTML-formaterede tekst, som
    indsættes øverst i den videresendte mail.

    mailbox kommer fra:
    item.data["box"]["mail"]["mailbox"]

    Mailadressen vises med fed skrift.
    Mailadressen er ikke et link.

    Dynamiske værdier bliver HTML-escaped,
    så specialtegn ikke ødelægger HTML'en.
    """

    def html_value(value):
        """
        Gør en dynamisk værdi sikker til HTML.

        Linjeskift bevares som HTML-linjeskift.
        """

        text = escape(
            str(
                value
                or ""
            )
        )

        return (
            text
            .replace(
                "\r\n",
                "<br>"
            )
            .replace(
                "\r",
                "<br>"
            )
            .replace(
                "\n",
                "<br>"
            )
        )


    mailbox_text = html_value(
        mailbox
    )


    subject_points = int(
        rule_result.get(
            "subject_points",
            0,
        )
        or 0
    )

    subject_match = html_value(
        rule_result.get(
            "subject_match",
            "",
        )
    )


    message_points = int(
        rule_result.get(
            "message_points",
            0,
        )
        or 0
    )

    message_match = html_value(
        rule_result.get(
            "message_match",
            "",
        )
    )


    sender_points = int(
        rule_result.get(
            "sender_points",
            0,
        )
        or 0
    )

    sender_match = html_value(
        rule_result.get(
            "sender_match",
            "",
        )
    )


    file_name_points = int(
        rule_result.get(
            "file_name_points",
            0,
        )
        or 0
    )

    file_name_match = html_value(
        rule_result.get(
            "file_name_match",
            "",
        )
    )


    file_content_points = int(
        rule_result.get(
            "file_content_points",
            0,
        )
        or 0
    )

    file_content_match = html_value(
        rule_result.get(
            "file_content_match",
            "",
        )
    )


    cpr_points = int(
        rule_result.get(
            "cpr_points",
            0,
        )
        or 0
    )


    rule_number = html_value(
        rule_result.get(
            "rule_number",
            "",
        )
    )


    return f"""
<div style="
    font-family: Calibri, Arial, sans-serif;
    font-size: 11pt;
    line-height: 1.35;
    color: #000000;
">

    <p style="margin: 0 0 16px 0;">
        Denne mail er automatisk fordelt af Robotten.
        Mailen er videresendt, da den har opnået
        300 point eller mere.
    </p>

    <p style="margin: 0 0 16px 0;">
        Hvis mailen ikke er videresendt korrekt,
        så gør følgende:<br>
        Videresend <strong>altid</strong> til
        <strong>{mailbox_text}</strong>.
        Skriv gerne hvem du tror der skulle have
        modtaget mailen og hvorfor.
    </p>

    <p style="margin: 0;">
        <u>
            Robotten har fundet disse ord/sætninger
            og fordelt point:
        </u><br>
        Emne({subject_points} point): {subject_match}<br>
        Besked({message_points} point): {message_match}<br>
        Afsender({sender_points} point): {sender_match}<br>
        Fil navn({file_name_points} point): {file_name_match}<br>
        Fil indhold({file_content_points} point):
        {file_content_match}<br>
        CPR fundet({cpr_points} point):<br>
        Regelnr: {rule_number}
    </p>

</div>
""".strip()
# -------------------------------------------------
# BUILD COMPARISON MESSAGE
# -------------------------------------------------

def build_comparison_message(
    rule_result,
    action_result,
):
    """
    Bygger Message til Automation Server.

    Under 300 point:
        Point: 150

    Mindst 300 point:
        Nr: 1035 + byg@haderslev.dk
    """

    total_points = int(
        rule_result.get(
            "total_points",
            0,
        )
        or 0
    )

    rule_number = rule_result.get(
        "rule_number"
    )

    if total_points < 300:
        return (
            f"Point: {total_points}"
        )

    if rule_number is None:
        return (
            f"Point: {total_points}"
        )

    message = (
        f"Nr: {rule_number}"
    )

    destination = (
        action_result.get("destination")
        or rule_result.get("destination")
        or ""
    )

    if destination:
        message += (
            f" + {destination}"
        )

    return message


# -------------------------------------------------
# UDFØR MAILHANDLING
# -------------------------------------------------

def execute_mail_action(
    mail,
    rule_result,
):
    """
    Udfører den endelige Outlook-handling.

    mail skal indeholde:
    - mailbox
    - message_id

    Regelvalg:

    Under 300 point:
        Overskriv kategorier.

    Testregel:
        Overskriv kategorier.

    Aktiv regel med mindst 300 point:
        Videresend eller flyt.
    """

    mailbox = mail[
        "mailbox"
    ]

    message_id = mail[
        "message_id"
    ]

    total_points = int(
        rule_result.get(
            "total_points",
            0,
        )
        or 0
    )

    rule_number = rule_result.get(
        "rule_number"
    )

    rule_status = (
        rule_result.get(
            "rule_status"
        )
        or ""
    ).strip().casefold()

    destination = (
        rule_result.get(
            "destination"
        )
        or ""
    ).strip()

    # -------------------------------------------------
    # BEDSTE REGEL UNDER 300 POINT
    # -------------------------------------------------

    if total_points < 300:
        categories = build_rule_categories(
            rule_result=rule_result,
            include_test=False,
        )

        update_mail_categories(
            user_mail=mailbox,
            message_id=message_id,
            categories=categories,
        )

        return {
            "action": "category_only",
            "destination": None,
            "categories": categories,
            "status_message": (
                "Kategorier opdateret. "
                "Mailen blev ikke fordelt"
            ),
            "simulated": False,
        }

    # -------------------------------------------------
    # TESTREGEL
    # -------------------------------------------------

    if rule_status == "test":
        categories = build_rule_categories(
            rule_result=rule_result,
            include_test=True,
        )

        update_mail_categories(
            user_mail=mailbox,
            message_id=message_id,
            categories=categories,
        )

        return {
            "action": "test_category",
            "destination": destination or None,
            "categories": categories,
            "status_message": (
                "Testkategorier opdateret"
            ),
            "simulated": False,
        }

    # -------------------------------------------------
    # AKTIV REGEL UDEN REGELNUMMER
    # -------------------------------------------------

    if rule_number is None:
        categories = build_rule_categories(
            rule_result=rule_result,
            include_test=False,
        )

        update_mail_categories(
            user_mail=mailbox,
            message_id=message_id,
            categories=categories,
        )

        return {
            "action": "category_only",
            "destination": None,
            "categories": categories,
            "status_message": (
                "Ingen regel valgt. "
                "Kategorier opdateret"
            ),
            "simulated": False,
        }

    # -------------------------------------------------
    # VIDERESEND TIL MAILADRESSE
    # -------------------------------------------------

    if destination_is_email(
        destination
    ):
        explanation = build_rule_explanation(
            rule_result=rule_result,
            mailbox=mailbox,
        )

        # -------------------------------------------------
        # 1. VIDERESEND MAILEN
        # -------------------------------------------------

        forward_result = forward_mail(
            user_mail=mailbox,
            message_id=message_id,
            forward={
                "to": [
                    destination
                ],
                "cc": [],
                "bcc": [],
                "body": explanation,
                "forward_mode": "formatted",
            },
        )


        # -------------------------------------------------
        # 2. SLET ORIGINALMAILEN
        # -------------------------------------------------

        delete_result = delete_mail(
            user_mail=mailbox,
            message_id=message_id,
            permanent_delete=False,
        )


        # -------------------------------------------------
        # 3. RETURNÉR RESULTAT
        # -------------------------------------------------

        return {
            "action": "forward_and_delete",
            "destination": destination,
            "categories": [],
            "status_message": (
                "Mail videresendt og flyttet "
                "til Slettet post"
            ),
            "explanation": explanation,
            "forward_result": forward_result,
            "delete_result": delete_result,
            "simulated": False,
        }

    # -------------------------------------------------
    # FLYT TIL OUTLOOK-MAPPE
    # -------------------------------------------------

    if destination:
        folder_id = find_folder_id(
            mailbox=mailbox,
            folder_name=destination,
        )

        moved_mail = move_mail(
            user_mail=mailbox,
            message_id=message_id,
            destination_folder_id=folder_id,
        )

        return {
            "action": "move",
            "destination": destination,
            "destination_folder_id": folder_id,
            "moved_message_id": (
                moved_mail.get("message_id")
                if moved_mail
                else None
            ),
            "categories": [],
            "status_message": (
                "Mail flyttet til undermappe"
            ),
            "simulated": False,
        }

    # -------------------------------------------------
    # REGEL UDEN DESTINATION
    # -------------------------------------------------

    categories = build_rule_categories(
        rule_result=rule_result,
        include_test=False,
    )

    update_mail_categories(
        user_mail=mailbox,
        message_id=message_id,
        categories=categories,
    )

    return {
        "action": "category_only",
        "destination": None,
        "categories": categories,
        "status_message": (
            "Regel nåede 300 point, "
            "men destinationen manglede. "
            "Kategorier opdateret"
        ),
        "simulated": False,
    }