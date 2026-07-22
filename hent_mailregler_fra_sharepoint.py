"""
HENT MAILREGLER FRA SHAREPOINT

Denne fil:

1. Henter alle konfigurerede Excel-regelark fra SharePoint.
2. Læser Excel-filerne direkte fra hukommelsen.
3. Finder automatisk den detaljerede overskriftsrække.
4. Gør dublerede Excel-kolonnenavne unikke.
5. Konverterer Excel-rækker til MailRule-objekter.
6. Fjerner regler med status Inaktiv.
7. Tilknytter hvert regelark til en postkasse.
8. Filtrerer reglerne til det aktuelle queue-item.
9. Understøtter RULE_MAILBOX_OVERRIDE fra .env.
10. Forsøger filhentning igen ved midlertidige netværksfejl.

Alle regelark hentes én gang ved workerens opstart.
Reglerne hentes ikke igen for hvert queue-item.
"""

import logging
import time
from dataclasses import asdict, dataclass
from typing import Any

import requests

from q_excel.excel_via_memory import (
    read_file_from_memory_as_dicts,
)

from q_sharepoint_api.sp_api import get_client

from proces_konfiguration import (
    RULE_MAILBOX_OVERRIDE,
)


logger = logging.getLogger(__name__)


# -------------------------------------------------
# DOWNLOADINDSTILLINGER
# -------------------------------------------------

SHAREPOINT_DOWNLOAD_MAX_ATTEMPTS = 3

SHAREPOINT_DOWNLOAD_WAIT_SECONDS = 5


# -------------------------------------------------
# EXCEL-KILDER
# -------------------------------------------------

@dataclass(frozen=True)
class RuleSource:
    """
    Dataclass for ét Excel-regelark.

    mailbox:
        Den postkasse reglerne tilhører.

    site_name:
        SharePoint-sitets navn.

    file_path:
        Filsti relativt fra dokumentbibliotekets rod.

    sheet_name:
        Excel-arkets navn eller indeks.
    """

    mailbox: str
    site_name: str
    file_path: str
    sheet_name: str | int = "Regler"


RULE_SOURCES = [
    RuleSource(
        mailbox="jobcenter@haderslev.dk",
        site_name="Automatisering",
        file_path=(
            "RPA - Processer/"
            "Fordeling af mails/"
            "Fællespostkasse - regler til emails.xlsx"
        ),
        sheet_name="Regler",
    ),

    # Tilføj flere regelark her senere.
    #
    # Eksempel:
    #
    # RuleSource(
    #     mailbox="post@haderslev.dk",
    #     site_name="Automatisering",
    #     file_path=(
    #         "RPA - Processer/"
    #         "Fordeling af mails/"
    #         "Regler til post.xlsx"
    #     ),
    #     sheet_name="Regler",
    # ),
]


# -------------------------------------------------
# REGELMODEL
# -------------------------------------------------

@dataclass(frozen=True)
class MailRule:
    """
    Fast datastruktur for én mailregel.
    """

    mailbox: str
    source_file: str
    excel_row_number: int

    rule_number: int
    status: str
    rule_type: str
    destination: str

    subject_pattern: str
    subject_points: int

    message_pattern: str
    message_points: int

    cpr_points: int

    file_name_pattern: str
    file_name_points: int

    file_content_pattern: str
    file_content_points: int

    sender_pattern: str
    sender_points: int

    max_points: int


# -------------------------------------------------
# KONVERTER VÆRDI TIL TEKST
# -------------------------------------------------

def _to_text(value):
    """
    Konverterer en Excel-værdi til renset tekst.

    None, NaN og teksten None bliver tom tekst.
    """

    if value is None:
        return ""

    text = str(value).strip()

    if text.casefold() in {
        "nan",
        "none",
    }:
        return ""

    return text


# -------------------------------------------------
# KONVERTER VÆRDI TIL HELTAL
# -------------------------------------------------

def _to_integer(
    value,
    default=0,
):
    """
    Konverterer en Excel-værdi til heltal.

    Eksempler:

    blank celle bliver 0
    300 bliver 300
    150.0 bliver 150
    -1000 bliver -1000

    Ugyldig tekst bliver standardværdien.
    """

    text = _to_text(value)

    if not text:
        return default

    text = text.replace(
        ",",
        ".",
    )

    try:
        return int(float(text))

    except (TypeError, ValueError):
        return default


# -------------------------------------------------
# NORMALISÉR KOLONNENAVN
# -------------------------------------------------

def _normalize_header(value):
    """
    Normaliserer et Excel-kolonnenavn.

    Store og små bogstaver ignoreres.
    Flere mellemrum samles til ét.
    """

    return " ".join(
        _to_text(value).split()
    ).casefold()


# -------------------------------------------------
# HENT RÆKKEVÆRDIER I KOLONNERÆKKEFØLGE
# -------------------------------------------------

def _get_row_values(row):
    """
    Returnerer rækkeværdier i kolonnerækkefølge.

    Når q-excel læser med has_header=False,
    er dictionary-nøglerne normalt kolonnenumre.
    """

    def sort_key(key):
        try:
            return 0, int(key)

        except (TypeError, ValueError):
            return 1, str(key)

    sorted_keys = sorted(
        row.keys(),
        key=sort_key,
    )

    return [
        row.get(key)
        for key in sorted_keys
    ]


# -------------------------------------------------
# ER RÆKKEN DEN RIGTIGE OVERSKRIFTSRÆKKE?
# -------------------------------------------------

def _is_header_row(values):
    """
    Kontrollerer om rækken indeholder
    regelarkets detaljerede kolonnenavne.
    """

    normalized_values = {
        _normalize_header(value)
        for value in values
        if _to_text(value)
    }

    required_headers = {
        "regelnr",
        "status",
        "regeltype",
        "point0",
        "point1",
        "point2",
        "point3",
        "point4",
        "point5",
    }

    return required_headers.issubset(
        normalized_values
    )


# -------------------------------------------------
# GØR DUBLERede KOLONNENAVNE UNIKKE
# -------------------------------------------------

def _make_unique_headers(values):
    """
    Gør dublerede kolonnenavne unikke.

    Eksempel:

    Fil navn
    Fil navn.1

    Fil indhold
    Fil indhold.1

    Afsender
    Afsender.1
    """

    headers = []
    name_counts = {}

    for column_number, value in enumerate(
        values,
        start=1,
    ):
        base_name = (
            _to_text(value)
            or f"Unnamed_{column_number}"
        )

        normalized_name = _normalize_header(
            base_name
        )

        current_count = name_counts.get(
            normalized_name,
            0,
        )

        if current_count == 0:
            unique_name = base_name
        else:
            unique_name = (
                f"{base_name}.{current_count}"
            )

        name_counts[normalized_name] = (
            current_count + 1
        )

        headers.append(unique_name)

    return headers


# -------------------------------------------------
# FIND DETALJERET OVERSKRIFTSRÆKKE
# -------------------------------------------------

def find_header_row(raw_rows):
    """
    Finder rækken med blandt andet:

    Regelnr
    Status
    Regeltype
    Point0 til Point5

    Returnerer:

    header_row_index:
        Rækkens indeks i raw_rows.

    headers:
        Liste med unikke kolonnenavne.
    """

    for row_index, row in enumerate(
        raw_rows
    ):
        values = _get_row_values(row)

        if _is_header_row(values):
            return (
                row_index,
                _make_unique_headers(values),
            )

    raise ValueError(
        "Kunne ikke finde regelarkets "
        "detaljerede overskriftsrække. "
        "Rækken skal blandt andet indeholde "
        "Regelnr, Status, Regeltype og "
        "Point0 til Point5."
    )


# -------------------------------------------------
# BYG NAVNGIVET EXCEL-RÆKKE
# -------------------------------------------------

def _build_named_row(
    raw_row,
    headers,
):
    """
    Kobler Excel-rækkens værdier til kolonnenavne.
    """

    values = _get_row_values(
        raw_row
    )

    named_row = {}

    for column_index, header in enumerate(
        headers
    ):
        if column_index < len(values):
            named_row[header] = values[
                column_index
            ]
        else:
            named_row[header] = None

    return named_row


# -------------------------------------------------
# FIND KOLONNEVÆRDI
# -------------------------------------------------

def _get_value(
    row,
    *column_names,
    default=None,
):
    """
    Finder den første eksisterende kolonne.

    Sammenligningen ignorerer store og små bogstaver.
    """

    normalized_row = {
        _normalize_header(key): value
        for key, value in row.items()
    }

    for column_name in column_names:
        normalized_name = _normalize_header(
            column_name
        )

        if normalized_name in normalized_row:
            return normalized_row[
                normalized_name
            ]

    return default


# -------------------------------------------------
# NORMALISÉR ÉN REGEL
# -------------------------------------------------

def _normalize_rule(
    row: dict[str, Any],
    mailbox: str,
    source_file: str,
    excel_row_number: int,
):
    """
    Konverterer én Excel-række til MailRule.

    Returnerer None, hvis rækken ikke har
    et gyldigt regelnummer.
    """

    rule_number = _to_integer(
        _get_value(
            row,
            "Regelnr",
            "Regel nr",
            "Regelnummer",
        )
    )

    if rule_number == 0:
        return None

    return MailRule(
        mailbox=mailbox,
        source_file=source_file,
        excel_row_number=excel_row_number,

        rule_number=rule_number,

        status=_to_text(
            _get_value(
                row,
                "Status",
            )
        ),

        rule_type=_to_text(
            _get_value(
                row,
                "Regeltype",
                "Regel type",
            )
        ),

        destination=_to_text(
            _get_value(
                row,
                "Send Til eller undermappe",
                "Send til eller undermappe",
                "Send Til",
                "Send til",
            )
        ),

        subject_pattern=_to_text(
            _get_value(
                row,
                "Emne",
            )
        ),

        subject_points=_to_integer(
            _get_value(
                row,
                "Point0",
            )
        ),

        message_pattern=_to_text(
            _get_value(
                row,
                "Besked",
            )
        ),

        message_points=_to_integer(
            _get_value(
                row,
                "Point1",
            )
        ),

        cpr_points=_to_integer(
            _get_value(
                row,
                "Point2",
            )
        ),

        file_name_pattern=_to_text(
            _get_value(
                row,
                "Fil navn.1",
                "Filnavn.1",
                "Fil navn",
                "Filnavn",
            )
        ),

        file_name_points=_to_integer(
            _get_value(
                row,
                "Point3",
            )
        ),

        file_content_pattern=_to_text(
            _get_value(
                row,
                "Fil indhold.1",
                "Filindhold.1",
                "Fil indhold",
                "Filindhold",
            )
        ),

        file_content_points=_to_integer(
            _get_value(
                row,
                "Point4",
            )
        ),

        sender_pattern=_to_text(
            _get_value(
                row,
                "Afsender.1",
                "Afsender",
            )
        ),

        sender_points=_to_integer(
            _get_value(
                row,
                "Point5",
            )
        ),

        max_points=_to_integer(
            _get_value(
                row,
                "Max Point",
                "Max point",
                "Maks point",
            )
        ),
    )


# -------------------------------------------------
# DOWNLOAD REGELFIL MED GENFORSØG
# -------------------------------------------------

def _download_rule_file_with_retry(
    client,
    site_id,
    source,
):
    """
    Henter én regelfil med kontrollerede genforsøg.

    Kun midlertidige forbindelsesfejl og timeout
    forsøges igen.

    Fejl i filsti, dataformat eller programkode
    forsøges ikke skjult gentaget.
    """

    last_error = None

    for attempt_number in range(
        1,
        SHAREPOINT_DOWNLOAD_MAX_ATTEMPTS + 1,
    ):
        try:
            logger.info(
                (
                    "Henter regelfil for %s. "
                    "Forsøg %s af %s."
                ),
                source.mailbox,
                attempt_number,
                SHAREPOINT_DOWNLOAD_MAX_ATTEMPTS,
            )

            result = (
                client.download_file_to_memory_by_path(
                    site_id=site_id,
                    file_path=source.file_path,
                    save_dir=None,
                )
            )

            if not isinstance(result, dict):
                raise TypeError(
                    "SharePoint-download returnerede "
                    "ikke en dictionary."
                )

            filename = result.get(
                "filename"
            )

            file_bytes = result.get(
                "file_bytes"
            )

            if not filename:
                raise ValueError(
                    "SharePoint-download mangler filename."
                )

            if not isinstance(
                file_bytes,
                bytes,
            ):
                raise ValueError(
                    "SharePoint-download mangler "
                    "file_bytes som bytes."
                )

            if len(file_bytes) == 0:
                raise ValueError(
                    "Den hentede Excel-fil er tom."
                )

            logger.info(
                (
                    "Regelfilen blev hentet: %s. "
                    "Størrelse: %s bytes."
                ),
                filename,
                len(file_bytes),
            )

            return result

        except (
            requests.exceptions.ConnectionError,
            requests.exceptions.Timeout,
        ) as error:
            last_error = error

            logger.warning(
                (
                    "Midlertidig forbindelsesfejl ved "
                    "hentning af regelfilen for %s. "
                    "Forsøg %s af %s. Fejl: %s"
                ),
                source.mailbox,
                attempt_number,
                SHAREPOINT_DOWNLOAD_MAX_ATTEMPTS,
                error,
            )

            if (
                attempt_number
                < SHAREPOINT_DOWNLOAD_MAX_ATTEMPTS
            ):
                time.sleep(
                    SHAREPOINT_DOWNLOAD_WAIT_SECONDS
                )

    raise ConnectionError(
        (
            "Regelfilen kunne ikke hentes fra "
            "SharePoint efter "
            f"{SHAREPOINT_DOWNLOAD_MAX_ATTEMPTS} "
            "forsøg. "
            f"Postkasse: {source.mailbox}. "
            f"Site: {source.site_name}. "
            f"Filsti: {source.file_path}. "
            f"Seneste fejl: {last_error}"
        )
    ) from last_error


# -------------------------------------------------
# DOWNLOAD REGELFIL TIL TEST
# -------------------------------------------------

def download_rules_file_from_sharepoint(
    source=None,
):
    """
    Henter en regelfil fra SharePoint.

    Funktionen bevares som offentlig testfunktion.

    Hvis source ikke udfyldes,
    bruges den første RuleSource.
    """

    if source is None:
        if not RULE_SOURCES:
            raise ValueError(
                "RULE_SOURCES er tom."
            )

        source = RULE_SOURCES[0]

    client = get_client()

    logger.info(
        "Finder SharePoint-site: %s",
        source.site_name,
    )

    site_id = client.get_site_id(
        source.site_name
    )

    return _download_rule_file_with_retry(
        client=client,
        site_id=site_id,
        source=source,
    )


# -------------------------------------------------
# LÆS RÅ EXCEL-RÆKKER
# -------------------------------------------------

def read_raw_rule_rows(
    file_result,
    sheet_name="Regler",
):
    """
    Læser Excel-filen uden fast overskriftsrække.

    Regelarket indeholder flere rækker før
    den detaljerede kolonneoverskrift.
    """

    return read_file_from_memory_as_dicts(
        filename=file_result["filename"],
        file_bytes=file_result["file_bytes"],
        has_header=False,
        sheet_name=sheet_name,
    )


# -------------------------------------------------
# KLARGØR REGLER FRA RÅ RÆKKER
# -------------------------------------------------

def _prepare_rules_from_rows(
    raw_rows,
    source,
    source_file,
):
    """
    Konverterer rå Excel-rækker til MailRule-objekter.

    Inaktive regler filtreres fra.
    Excel-rækkefølgen bevares.
    """

    if not raw_rows:
        raise ValueError(
            f"Regelfilen '{source_file}' "
            "indeholder ingen rækker."
        )

    header_index, headers = (
        find_header_row(raw_rows)
    )

    rules = []

    data_rows = raw_rows[
        header_index + 1:
    ]

    for offset, raw_row in enumerate(
        data_rows,
        start=1,
    ):
        named_row = _build_named_row(
            raw_row=raw_row,
            headers=headers,
        )

        excel_row_number = (
            header_index
            + offset
            + 1
        )

        rule = _normalize_rule(
            row=named_row,
            mailbox=source.mailbox,
            source_file=source_file,
            excel_row_number=excel_row_number,
        )

        if rule is None:
            continue

        if (
            rule.status.casefold()
            == "inaktiv"
        ):
            continue

        rules.append(rule)

    return rules


# -------------------------------------------------
# HENT ÉT REGELARK
# -------------------------------------------------

def _load_rule_source(source):
    """
    Henter og læser ét Excel-regelark.
    """

    file_result = (
        download_rules_file_from_sharepoint(
            source=source
        )
    )

    filename = file_result[
        "filename"
    ]

    raw_rows = read_raw_rule_rows(
        file_result=file_result,
        sheet_name=source.sheet_name,
    )

    rules = _prepare_rules_from_rows(
        raw_rows=raw_rows,
        source=source,
        source_file=filename,
    )

    logger.info(
        "%s regler blev indlæst fra %s for %s.",
        len(rules),
        filename,
        source.mailbox,
    )

    return rules


# -------------------------------------------------
# HENT ALLE REGLER
# -------------------------------------------------

def get_all_mail_rules():
    """
    Henter regler fra alle konfigurerede Excel-filer.

    Funktionen skal kaldes én gang ved
    workerens opstart.
    """

    if not RULE_SOURCES:
        raise ValueError(
            "RULE_SOURCES indeholder ingen regelark."
        )

    all_rules = []

    for source in RULE_SOURCES:
        source_rules = _load_rule_source(
            source
        )

        all_rules.extend(
            source_rules
        )

    if not all_rules:
        raise ValueError(
            "Der blev ikke fundet nogen aktive "
            "eller testregler."
        )

    logger.info(
        "%s regler blev samlet indlæst.",
        len(all_rules),
    )

    return all_rules


# -------------------------------------------------
# FILTRÉR REGLER TIL POSTKASSE
# -------------------------------------------------

def get_rules_for_mailbox(
    all_rules,
    item_mailbox,
):
    """
    Filtrerer allerede indlæste regler til itemet.

    RULE_MAILBOX_OVERRIDE bruges, hvis værdien
    er udfyldt i .env.

    Eksempel:

    TEST_MAILBOX_OVERRIDE=robot-data@haderslev.dk
    RULE_MAILBOX_OVERRIDE=jobcenter@haderslev.dk

    Mailen hentes da fra robot-data, men vurderes
    med reglerne for jobcenter.
    """

    mailbox_to_use = (
        RULE_MAILBOX_OVERRIDE
        or item_mailbox
    )

    normalized_mailbox = (
        str(mailbox_to_use)
        .strip()
        .casefold()
    )

    filtered_rules = [
        rule
        for rule in all_rules
        if (
            rule.mailbox
            .strip()
            .casefold()
            == normalized_mailbox
        )
    ]

    if not filtered_rules:
        available_mailboxes = sorted(
            {
                rule.mailbox
                for rule in all_rules
            }
        )

        raise ValueError(
            "Ingen regler blev fundet for "
            f"postkassen '{mailbox_to_use}'. "
            "Tilgængelige regelpostkasser: "
            f"{available_mailboxes}. "
            "Kontrollér RULE_MAILBOX_OVERRIDE "
            "i .env."
        )

    logger.info(
        (
            "%s regler blev valgt for "
            "item-postkassen %s. "
            "Regelpostkasse: %s."
        ),
        len(filtered_rules),
        item_mailbox,
        mailbox_to_use,
    )

    return filtered_rules


# -------------------------------------------------
# BAGUDKOMPATIBEL TESTFUNKTION
# -------------------------------------------------

def get_mail_rules():
    """
    Henter regler fra den første Excel-kilde.

    Funktionen bevares til den eksisterende testfil.

    Worker-processen bør bruge:
    get_all_mail_rules()
    """

    if not RULE_SOURCES:
        raise ValueError(
            "RULE_SOURCES er tom."
        )

    return _load_rule_source(
        RULE_SOURCES[0]
    )


# -------------------------------------------------
# REGEL TIL DICTIONARY
# -------------------------------------------------

def rule_to_dict(rule):
    """
    Konverterer MailRule til dictionary.

    Funktionen bruges i test og fejlsøgning.
    """

    return asdict(rule)