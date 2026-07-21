"""
HENT MAILREGLER FRA SHAREPOINT

Denne fil:

1. Henter Excel-filen fra SharePoint som bytes.
2. Læser alle Excel-rækker uden at antage en fast overskriftsrække.
3. Finder automatisk rækken med de rigtige kolonnenavne.
4. Håndterer dublerede kolonnenavne.
5. Konverterer hver Excel-række til en MailRule.
6. Fjerner kun regler med status Inaktiv.
7. Bevarer rækkefølgen fra Excel.

Reglerne skal hentes én gang ved workerens opstart.
Reglerne skal ikke hentes i --queue.
"""

from dataclasses import asdict, dataclass
from typing import Any

from q_excel.excel_via_memory import (
    read_file_from_memory_as_dicts,
)

from q_sharepoint_api.sp_api import get_client

from proces_konfiguration import (
    RULES_FILE_PATH,
    RULES_SHEET_NAME,
    RULES_SITE_NAME,
    RULE_STATUS_INACTIVE,
)


# -------------------------------------------------
# REGELMODEL
# -------------------------------------------------

@dataclass(frozen=True)
class MailRule:
    """
    Dataclass (fast datastruktur) for én mailregel.
    """

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
# HENT EXCEL-FIL FRA SHAREPOINT
# -------------------------------------------------

def download_rules_file_from_sharepoint():
    """
    Henter regelfilen fra SharePoint til hukommelsen.

    Returnerer dictionary (nøgle-værdi-samling):

    {
        "filename": "...xlsx",
        "file_bytes": b"..."
    }
    """

    client = get_client()

    site_id = client.get_site_id(
        RULES_SITE_NAME
    )

    result = client.download_file_to_memory_by_path(
        site_id=site_id,
        file_path=RULES_FILE_PATH,
        save_dir=None,
    )

    if not isinstance(result, dict):
        raise TypeError(
            "SharePoint-download returnerede ikke en dictionary."
        )

    filename = result.get("filename")
    file_bytes = result.get("file_bytes")

    if not filename:
        raise ValueError(
            "SharePoint-download mangler filename."
        )

    if not isinstance(file_bytes, bytes):
        raise ValueError(
            "SharePoint-download mangler file_bytes som bytes."
        )

    if len(file_bytes) == 0:
        raise ValueError(
            "SharePoint-filen er tom."
        )

    return result


# -------------------------------------------------
# KONVERTER VÆRDI TIL TEKST
# -------------------------------------------------

def _to_text(value):
    """
    Konverterer værdi til renset tekst.
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
    Konverterer værdi til int (heltal).
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

def _normalize_column_name(value):
    """
    Normaliserer et kolonnenavn til sammenligning.
    """

    text = _to_text(value)

    return " ".join(
        text.split()
    ).casefold()


# -------------------------------------------------
# FIND SORTEREDE RÆKKEVÆRDIER
# -------------------------------------------------

def _get_row_values(row):
    """
    Returnerer rækkeværdier i kolonnerækkefølge.

    Når Excel læses med has_header=False, er nøglerne
    normalt kolonnenumre.
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
# ER DET DEN RIGTIGE OVERSKRIFTSRÆKKE?
# -------------------------------------------------

def _is_header_row(values):
    """
    Kontrollerer om rækken er den detaljerede
    overskriftsrække fra regelarket.
    """

    normalized_values = {
        _normalize_column_name(value)
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
    """

    headers = []
    used_names = {}

    for column_index, value in enumerate(
        values,
        start=1,
    ):
        base_name = _to_text(value)

        if not base_name:
            base_name = (
                f"Unnamed_{column_index}"
            )

        normalized_name = (
            _normalize_column_name(
                base_name
            )
        )

        duplicate_number = used_names.get(
            normalized_name,
            0,
        )

        if duplicate_number == 0:
            unique_name = base_name
        else:
            unique_name = (
                f"{base_name}."
                f"{duplicate_number}"
            )

        used_names[normalized_name] = (
            duplicate_number + 1
        )

        headers.append(unique_name)

    return headers


# -------------------------------------------------
# FIND OVERSKRIFTSRÆKKEN
# -------------------------------------------------

def find_header_row(raw_rows):
    """
    Finder Excel-rækken med:

    Regelnr
    Status
    Regeltype
    Point0 til Point5

    Returnerer:
    - rækkens indeks i listen
    - unikke kolonnenavne
    """

    for row_index, raw_row in enumerate(
        raw_rows
    ):
        values = _get_row_values(
            raw_row
        )

        if _is_header_row(values):
            headers = _make_unique_headers(
                values
            )

            return row_index, headers

    raise ValueError(
        "Kunne ikke finde regelarkets overskriftsrække. "
        "Rækken skal blandt andet indeholde Regelnr, "
        "Status, Regeltype og Point0 til Point5."
    )


# -------------------------------------------------
# BYG DICTIONARY FRA DATA-RÆKKE
# -------------------------------------------------

def _build_named_row(
    raw_row,
    headers,
):
    """
    Kobler en datarække sammen med overskrifterne.
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
# FIND VÆRDI VED KOLONNENAVN
# -------------------------------------------------

def _get_value(
    row,
    *column_names,
    default=None,
):
    """
    Finder den første eksisterende kolonne.
    """

    normalized_row = {
        _normalize_column_name(key): value
        for key, value in row.items()
    }

    for column_name in column_names:
        normalized_name = (
            _normalize_column_name(
                column_name
            )
        )

        if normalized_name in normalized_row:
            return normalized_row[
                normalized_name
            ]

    return default


# -------------------------------------------------
# NORMALISÉR ÉN REGEL
# -------------------------------------------------

def normalize_rule(
    row: dict[str, Any],
    excel_row_number: int,
):
    """
    Konverterer én Excel-række til MailRule.

    Returnerer None, hvis rækken ikke indeholder
    et regelnummer.
    """

    rule_number = _to_integer(
        _get_value(
            row,
            "Regelnr",
            "Regel nr",
            "Regelnummer",
        ),
        default=0,
    )

    if rule_number == 0:
        return None

    return MailRule(
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
                "Fil navn",
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
                "Fil indhold",
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
            ),
            default=300,
        ),
    )


# -------------------------------------------------
# LÆS RÅ EXCEL-RÆKKER
# -------------------------------------------------

def read_raw_rule_rows(file_result):
    """
    Læser Excel-filen uden fast overskriftsrække.

    Det er nødvendigt, fordi arket indeholder:
    - titel
    - vejledning
    - overordnet kolonnerække
    - detaljeret kolonnerække
    """

    return read_file_from_memory_as_dicts(
        filename=file_result["filename"],
        file_bytes=file_result["file_bytes"],
        has_header=False,
        sheet_name=RULES_SHEET_NAME,
    )


# -------------------------------------------------
# KLARGØR REGLER
# -------------------------------------------------

def prepare_rules(raw_rows):
    """
    Finder overskriftsrækken og konverterer
    alle efterfølgende datarækker til regler.
    """

    header_row_index, headers = (
        find_header_row(raw_rows)
    )

    rules = []

    data_rows = raw_rows[
        header_row_index + 1:
    ]

    for list_index, raw_row in enumerate(
        data_rows,
        start=1,
    ):
        excel_row_number = (
            header_row_index
            + list_index
            + 1
        )

        named_row = _build_named_row(
            raw_row=raw_row,
            headers=headers,
        )

        rule = normalize_rule(
            row=named_row,
            excel_row_number=(
                excel_row_number
            ),
        )

        if rule is None:
            continue

        if (
            rule.status.casefold()
            == RULE_STATUS_INACTIVE.casefold()
        ):
            continue

        rules.append(rule)

    return rules


# -------------------------------------------------
# HENT MAILREGLER
# -------------------------------------------------

def get_mail_rules():
    """
    Henter og klargør alle aktive og testregler.

    Funktionen kaldes én gang ved workerens opstart.
    """

    file_result = (
        download_rules_file_from_sharepoint()
    )

    raw_rows = read_raw_rule_rows(
        file_result
    )

    rules = prepare_rules(
        raw_rows
    )

    if not rules:
        raise ValueError(
            "Der blev ikke fundet aktive eller "
            "testregler i Excel-filen."
        )

    return rules


# -------------------------------------------------
# REGEL TIL DICTIONARY
# -------------------------------------------------

def rule_to_dict(rule):
    """
    Konverterer MailRule til dictionary.

    Funktionen bruges blandt andet i testen.
    """

    return asdict(rule)