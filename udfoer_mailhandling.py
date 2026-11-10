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
from email.utils import parseaddr

from q_outlook_api.functionality.mail_api import (
    delete_mail,
    forward_mail,
    get_folders,
    move_mail,
    update_mail_categories,
)

from proces_konfiguration import (
    ALLOWED_FORWARD_DOMAINS,
)


# -------------------------------------------------
# SER DESTINATIONEN UD SOM MAILADRESSER?
# -------------------------------------------------
def destination_contains_email(destination):
    """
    Kontrollerer, om destinationen skal behandles som en
    maildestination frem for en Outlook-mappe.

    Returnerer:
        True:
            Destinationen indeholder mindst ét @.
        False:
            Destinationen indeholder ikke @.
    """
    destination_text = str(
        destination or ""
    ).strip()

    return "@" in destination_text


# -------------------------------------------------
# SPLIT MAILDESTINATIONER
# -------------------------------------------------
def split_email_destinations(destination):
    """
    Splitter en semikolonsepareret maildestination.

    Eksempel:
        Input:
            "a@haderslev.dk; b@haderslev.dk"

        Output:
            [
                "a@haderslev.dk",
                "b@haderslev.dk",
            ]

    Der fjernes mellemrum omkring adresserne.

    En blank destination returnerer en tom liste.

    Tomme elementer mellem semikoloner accepteres ikke.
    Det forhindrer eksempelvis:
        "a@haderslev.dk;;b@haderslev.dk"
    """
    destination_text = str(
        destination or ""
    ).strip()

    if not destination_text:
        return []

    raw_destinations = destination_text.split(";")

    email_destinations = []

    for position, raw_destination in enumerate(
        raw_destinations,
        start=1,
    ):
        email_destination = raw_destination.strip()

        if not email_destination:
            raise ValueError(
                "Maildestinationen indeholder en tom "
                f"adresse på position {position}. "
                "Kontrollér placeringen af semikoloner."
            )

        email_destinations.append(
            email_destination
        )

    return email_destinations


# -------------------------------------------------
# ER ÉN DESTINATION EN GYLDIG MAILADRESSE?
# -------------------------------------------------
def destination_is_email(destination):
    """
    Kontrollerer én enkelt mailadresse.

    Funktionen accepterer ikke flere adresser i samme tekst.
    Flere adresser skal først splittes med
    split_email_destinations().

    Returnerer:
        True:
            Adressen har et gyldigt grundformat.
        False:
            Adressen har et ugyldigt format.
    """
    destination_text = str(
        destination or ""
    ).strip()

    if not destination_text:
        return False

    # Semikolon må kun bruges til at adskille adresser
    # før denne funktion kaldes.
    if ";" in destination_text:
        return False

    # Mellemrum accepteres ikke inde i en mailadresse.
    if any(
        character.isspace()
        for character in destination_text
    ):
        return False

    display_name, parsed_address = parseaddr(
        destination_text
    )

    # Visningsnavne som:
    # "Rune Hedegaard <rune@haderslev.dk>"
    # accepteres ikke i Excel-feltet.
    if display_name:
        return False

    # parseaddr må ikke ændre den indtastede adresse.
    if parsed_address != destination_text:
        return False

    if parsed_address.count("@") != 1:
        return False

    local_part, domain = parsed_address.rsplit(
        "@",
        1,
    )

    if not local_part or not domain:
        return False

    if local_part.startswith("."):
        return False

    if local_part.endswith("."):
        return False

    if ".." in local_part:
        return False

    if domain.startswith("."):
        return False

    if domain.endswith("."):
        return False

    if "." not in domain:
        return False

    domain_parts = domain.split(".")

    if any(
        not domain_part
        for domain_part in domain_parts
    ):
        return False

    if any(
        domain_part.startswith("-")
        or domain_part.endswith("-")
        for domain_part in domain_parts
    ):
        return False

    return True


# -------------------------------------------------
# HENT DOMÆNET FRA ÉN MAILADRESSE
# -------------------------------------------------
def get_destination_domain(destination):
    """
    Henter domænet fra én valideret mailadresse.

    Eksempel:
        Input:
            medarbejder@haderslev.dk

        Output:
            haderslev.dk

    Hvis destinationen ikke er én gyldig mailadresse,
    returneres tom tekst.
    """
    destination_text = str(
        destination or ""
    ).strip()

    if not destination_is_email(
        destination_text
    ):
        return ""

    _, domain = destination_text.rsplit(
        "@",
        1,
    )

    return (
        domain
        .strip()
        .rstrip(".")
        .casefold()
    )


# -------------------------------------------------
# ER ÉN MAILADRESSE TILLADT?
# -------------------------------------------------
def destination_email_is_allowed(destination):
    """
    Kontrollerer om én mailadresses domæne findes
    i ALLOWED_FORWARD_DOMAINS.

    Sammenligningen:
    - ignorerer store og små bogstaver
    - fjerner mellemrum
    - kræver et præcist domænematch

    Funktionen skal kaldes én gang pr. adresse.
    """
    destination_domain = get_destination_domain(
        destination
    )

    if not destination_domain:
        return False

    allowed_domains = {
        str(domain)
        .strip()
        .rstrip(".")
        .casefold()
        for domain in ALLOWED_FORWARD_DOMAINS
        if str(domain).strip()
    }

    return destination_domain in allowed_domains


# -------------------------------------------------
# VALIDÉR ALLE MAILDESTINATIONER
# -------------------------------------------------
def validate_email_destinations(destination):
    """
    Splitter og validerer alle mailadresser.

    Alle adresser skal være:
    - udfyldt
    - gyldige enkeltadresser
    - på et tilladt domæne

    Returnerer:
        Dictionary med:
        - valid
        - destinations
        - invalid_destinations
        - blocked_destinations
        - error_message

    Der udføres ingen Outlook-handling i denne funktion.
    """
    try:
        destinations = split_email_destinations(
            destination
        )
    except ValueError as error:
        return {
            "valid": False,
            "destinations": [],
            "invalid_destinations": [],
            "blocked_destinations": [],
            "error_message": str(error),
        }

    if not destinations:
        return {
            "valid": False,
            "destinations": [],
            "invalid_destinations": [],
            "blocked_destinations": [],
            "error_message": (
                "Der blev ikke fundet nogen mailadresser."
            ),
        }

    invalid_destinations = [
        email_destination
        for email_destination in destinations
        if not destination_is_email(
            email_destination
        )
    ]

    if invalid_destinations:
        return {
            "valid": False,
            "destinations": destinations,
            "invalid_destinations": invalid_destinations,
            "blocked_destinations": [],
            "error_message": (
                "En eller flere mailadresser har "
                "et ugyldigt format."
            ),
        }

    blocked_destinations = [
        email_destination
        for email_destination in destinations
        if not destination_email_is_allowed(
            email_destination
        )
    ]

    if blocked_destinations:
        return {
            "valid": False,
            "destinations": destinations,
            "invalid_destinations": [],
            "blocked_destinations": blocked_destinations,
            "error_message": (
                "En eller flere mailadresser har "
                "et domæne, som ikke er tilladt."
            ),
        }

    return {
        "valid": True,
        "destinations": destinations,
        "invalid_destinations": [],
        "blocked_destinations": [],
        "error_message": None,
    }




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
        Validér eventuelle maildestinationer.
        Videresend ikke.

    Aktiv regel med mindst 300 point:
        Videresend til én eller flere mailadresser,
        eller flyt til en Outlook-mappe.

    Sikkerhed:
        Alle mailadresser valideres enkeltvis.
        Alle adresser skal tilhøre et domæne i
        ALLOWED_FORWARD_DOMAINS.

        Hvis blot én adresse er ugyldig eller blokeret:
        - videresendes mailen ikke
        - originalmailen slettes ikke
        - mailen får en fejlkategori
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
            "destinations": [],
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

        if destination:
            categories.append(
                f"Til: {destination}"
            )

        validation_result = None

        if destination_contains_email(
            destination
        ):
            validation_result = (
                validate_email_destinations(
                    destination
                )
            )

            if validation_result[
                "invalid_destinations"
            ]:
                categories.append(
                    "Ugyldig maildestination"
                )
            elif validation_result[
                "blocked_destinations"
            ]:
                categories.append(
                    "Ekstern destination blokeret"
                )

        update_mail_categories(
            user_mail=mailbox,
            message_id=message_id,
            categories=categories,
        )

        if (
            validation_result
            and validation_result[
                "invalid_destinations"
            ]
        ):
            status_message = (
                "Testkategorier opdateret. "
                "En eller flere mailadresser er ugyldige"
            )
        elif (
            validation_result
            and validation_result[
                "blocked_destinations"
            ]
        ):
            status_message = (
                "Testkategorier opdateret. "
                "Et eller flere domæner er ikke tilladt"
            )
        elif destination:
            status_message = (
                "Testkategorier opdateret. "
                "Destinationen blev fundet"
            )
        else:
            status_message = (
                "Testkategorier opdateret. "
                "Destinationen manglede"
            )

        return {
            "action": "test_category",
            "destination": destination or None,
            "destinations": (
                validation_result["destinations"]
                if validation_result
                else []
            ),
            "categories": categories,
            "status_message": status_message,
            "validation": validation_result,
            "simulated": True,
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
            "destinations": [],
            "categories": categories,
            "status_message": (
                "Ingen regel valgt. "
                "Kategorier opdateret"
            ),
            "simulated": False,
        }

    # -------------------------------------------------
    # DESTINATIONEN INDEHOLDER MAILADRESSER
    # -------------------------------------------------
    if destination_contains_email(
        destination
    ):
        validation_result = (
            validate_email_destinations(
                destination
            )
        )

        destinations = validation_result[
            "destinations"
        ]

        invalid_destinations = validation_result[
            "invalid_destinations"
        ]

        blocked_destinations = validation_result[
            "blocked_destinations"
        ]

        # ---------------------------------------------
        # UGYLDIGE MAILADRESSER
        # ---------------------------------------------
        if (
            not validation_result["valid"]
            and not blocked_destinations
        ):
            categories = build_rule_categories(
                rule_result=rule_result,
                include_test=False,
            )

            categories.append(
                "Ugyldig maildestination"
            )

            update_mail_categories(
                user_mail=mailbox,
                message_id=message_id,
                categories=categories,
            )

            return {
                "action": "invalid_email_destination",
                "destination": destination,
                "destinations": destinations,
                "invalid_destinations": (
                    invalid_destinations
                ),
                "blocked_destinations": [],
                "categories": categories,
                "status_message": (
                    "Mailen blev ikke videresendt. "
                    f"{validation_result['error_message']}"
                ),
                "simulated": False,
            }

        # ---------------------------------------------
        # BLOKEREDE MAILADRESSER
        # ---------------------------------------------
        if blocked_destinations:
            categories = build_rule_categories(
                rule_result=rule_result,
                include_test=False,
            )

            categories.append(
                "Ekstern destination blokeret"
            )

            update_mail_categories(
                user_mail=mailbox,
                message_id=message_id,
                categories=categories,
            )

            return {
                "action": (
                    "external_destination_blocked"
                ),
                "destination": destination,
                "destinations": destinations,
                "invalid_destinations": [],
                "blocked_destinations": (
                    blocked_destinations
                ),
                "categories": categories,
                "status_message": (
                    "Mailen blev ikke videresendt. "
                    "Et eller flere destinationsdomæner "
                    "er ikke tilladt"
                ),
                "simulated": False,
            }

        # ---------------------------------------------
        # ALLE ADRESSER ER GODKENDT
        # ---------------------------------------------
        explanation = build_rule_explanation(
            rule_result=rule_result,
            mailbox=mailbox,
        )

        # forward_mail kaldes først, når samtlige
        # adresser er valideret og godkendt.
        forward_result = forward_mail(
            user_mail=mailbox,
            message_id=message_id,
            forward={
                "to": destinations,
                "cc": [],
                "bcc": [],
                "body": explanation,
                "forward_mode": "formatted",
            },
        )

        # Originalmailen slettes kun, hvis forward_mail
        # ikke har rejst en fejl.
        delete_result = delete_mail(
            user_mail=mailbox,
            message_id=message_id,
            permanent_delete=False,
        )

        return {
            "action": "forward_and_delete",
            "destination": destination,
            "destinations": destinations,
            "categories": [],
            "status_message": (
                f"Mail videresendt til "
                f"{len(destinations)} modtager(e) og "
                "flyttet til Slettet post"
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
            "destinations": [],
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
        "destinations": [],
        "categories": categories,
        "status_message": (
            "Regel nåede 300 point, "
            "men destinationen manglede. "
            "Kategorier opdateret"
        ),
        "simulated": False,
    }