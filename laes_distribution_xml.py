"""
LÆS MODTAGER FRA DISTRIBUTION.XML

Denne fil håndterer dynamiske destinationer fra vedhæftninger.

Understøttet pladsholder i Excel:

    {{distribution.xml:internal_email}}

Når den valgte regel indeholder denne værdi, findes
vedhæftningen distribution.xml, hvorefter mailadressen
læses fra XML-elementet internal_email.

Vedhæftningsindhold og mailadresse gemmes ikke permanent
af denne funktion.
"""

from xml.etree import ElementTree


DISTRIBUTION_XML_PLACEHOLDER = (
    "{{distribution.xml:internal_email}}"
)

DISTRIBUTION_XML_FILENAME = "distribution.xml"

DISTRIBUTION_XML_EMAIL_ELEMENT = "internal_email"


def bruger_distribution_xml(destination):
    """
    Kontrollerer om destinationen i Excel betyder,
    at modtageren skal læses fra distribution.xml.

    Output:
        True:
            Destinationen er XML-pladsholderen.

        False:
            Destinationen skal behandles normalt.
    """
    normalized_destination = (
        str(destination or "")
        .strip()
        .casefold()
    )

    return (
        normalized_destination
        == DISTRIBUTION_XML_PLACEHOLDER.casefold()
    )


def _find_distribution_xml_attachment(attachments):
    """
    Finder distribution.xml blandt mailens vedhæftninger.

    Der kontrolleres både:
    - name
    - original_name

    original_name bruges også, fordi q-outlook-api kan gøre
    dublerede filnavne unikke, eksempelvis:

        distribution.xml
        distribution (1).xml

    Output:
        Dictionaryen for den fundne vedhæftning.

    Fejl:
        ValueError hvis filen mangler, eller hvis flere
        vedhæftninger oprindeligt hedder distribution.xml.
    """
    matches = []

    for attachment in attachments or []:
        current_name = (
            str(
                attachment.get("name")
                or ""
            )
            .strip()
            .casefold()
        )

        original_name = (
            str(
                attachment.get("original_name")
                or ""
            )
            .strip()
            .casefold()
        )

        if (
            current_name
            == DISTRIBUTION_XML_FILENAME.casefold()
            or original_name
            == DISTRIBUTION_XML_FILENAME.casefold()
        ):
            matches.append(attachment)

    if not matches:
        raise ValueError(
            "Reglen kræver en vedhæftning med navnet "
            f"'{DISTRIBUTION_XML_FILENAME}', men filen "
            "blev ikke fundet på mailen."
        )

    if len(matches) > 1:
        raise ValueError(
            "Mailen indeholder flere vedhæftninger med "
            f"navnet '{DISTRIBUTION_XML_FILENAME}'. "
            "Robotten kan derfor ikke afgøre, hvilken "
            "fil der skal bruges."
        )

    return matches[0]


def _get_local_xml_name(tag):
    """
    Fjerner et eventuelt XML-namespace fra elementnavnet.

    Eksempel:

        {urn:test}internal_email

    bliver til:

        internal_email
    """
    return str(tag).rsplit(
        "}",
        maxsplit=1,
    )[-1]


def _find_internal_email_element(root):
    """
    Finder internal_email i hele XML-træet.

    Funktionen understøtter både XML med og uden namespace.

    Output:
        Det fundne XML-element.

    Fejl:
        ValueError hvis elementet mangler eller findes
        flere gange.
    """
    matches = [
        element
        for element in root.iter()
        if (
            _get_local_xml_name(element.tag)
            .strip()
            .casefold()
            == DISTRIBUTION_XML_EMAIL_ELEMENT.casefold()
        )
    ]

    if not matches:
        raise ValueError(
            "Vedhæftningen distribution.xml indeholder ikke "
            "XML-elementet 'internal_email'."
        )

    if len(matches) > 1:
        raise ValueError(
            "Vedhæftningen distribution.xml indeholder flere "
            "XML-elementer med navnet 'internal_email'. "
            "Robotten kan derfor ikke vælge en modtager."
        )

    return matches[0]


def _validate_email_address(email_address):
    """
    Udfører en grundlæggende kontrol af mailadressen.

    Den eksisterende domænekontrol i udfoer_mailhandling.py
    skal stadig udføres bagefter.

    Output:
        Den rensede mailadresse.

    Fejl:
        ValueError hvis adressen ikke ligner én enkelt
        mailadresse.
    """
    cleaned_email = str(
        email_address or ""
    ).strip()

    if cleaned_email.count("@") != 1:
        raise ValueError(
            "internal_email i distribution.xml indeholder "
            "ikke én gyldig mailadresse."
        )

    local_part, domain = cleaned_email.rsplit(
        "@",
        1,
    )

    if (
        not local_part.strip()
        or not domain.strip()
        or "." not in domain
        or any(
            character.isspace()
            for character in cleaned_email
        )
    ):
        raise ValueError(
            "internal_email i distribution.xml indeholder "
            "ikke en gyldig mailadresse."
        )

    return cleaned_email


def laes_internal_email_fra_distribution_xml(
    attachments,
):
    """
    Finder distribution.xml og læser internal_email.

    Input:
        attachments:
            Listen fra q-outlook-api-funktionen
            get_attachments().

    Output:
        Mailadressen som renset tekst.

    Fejl:
        ValueError hvis:
        - distribution.xml mangler
        - flere distribution.xml-filer findes
        - filen mangler bytes
        - XML'en er ugyldig
        - internal_email mangler
        - mailadressen er ugyldig
    """
    attachment = (
        _find_distribution_xml_attachment(
            attachments
        )
    )

    xml_bytes = attachment.get(
        "content_bytes"
    )

    if not isinstance(xml_bytes, bytes):
        raise ValueError(
            "Vedhæftningen distribution.xml mangler "
            "content_bytes."
        )

    if not xml_bytes:
        raise ValueError(
            "Vedhæftningen distribution.xml er tom."
        )

    try:
        root = ElementTree.fromstring(
            xml_bytes
        )
    except ElementTree.ParseError as error:
        raise ValueError(
            "Vedhæftningen distribution.xml indeholder "
            "ugyldig XML."
        ) from error

    internal_email_element = (
        _find_internal_email_element(root)
    )

    return _validate_email_address(
        internal_email_element.text
    )


def resolve_rule_destination(
    rule_result,
    attachments,
):
    """
    Erstatter en dynamisk Excel-destination med
    internal_email fra distribution.xml.

    Funktionen ændrer en kopi af rule_result og ændrer
    derfor ikke den oprindelige dictionary direkte.

    Pointberegningen ændres ikke.

    Output:
        En kopi af rule_result med:
        - destination
        - destination_source

    Hvis destinationen ikke er XML-pladsholderen,
    returneres resultatet med den almindelige
    Excel-destination.
    """
    resolved_result = dict(
        rule_result or {}
    )

    excel_destination = (
        resolved_result.get("destination")
        or ""
    ).strip()

    if not bruger_distribution_xml(
        excel_destination
    ):
        resolved_result["destination_source"] = (
            "excel"
        )
        return resolved_result

    xml_email = (
        laes_internal_email_fra_distribution_xml(
            attachments
        )
    )

    resolved_result["destination"] = xml_email
    resolved_result["destination_source"] = (
        "distribution.xml"
    )

    return resolved_result