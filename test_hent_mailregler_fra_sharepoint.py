"""
TEST AF MAILREGLER FRA SHAREPOINT

Testen:

1. Henter Excel fra SharePoint.
2. Kontrollerer filename og bytes.
3. Finder den rigtige overskriftsrække.
4. Læser aktive og testregler.
5. Kontrollerer centrale felter.
6. Printer et overskueligt resultat.

Testen ændrer ikke SharePoint-filen.
Testen ændrer ikke mails.
Testen ændrer ikke Automation Server-items.
"""

from pprint import pprint

from hent_mailregler_fra_sharepoint import (
    download_rules_file_from_sharepoint,
    find_header_row,
    get_mail_rules,
    read_raw_rule_rows,
    rule_to_dict,
)


# -------------------------------------------------
# TEST SHAREPOINT-DOWNLOAD
# -------------------------------------------------

def test_download_rules_file():
    """
    Tester at Excel-filen kan hentes som bytes.
    """

    print(
        "\n1. Henter regelfilen fra SharePoint..."
    )

    result = (
        download_rules_file_from_sharepoint()
    )

    assert isinstance(
        result,
        dict,
    ), "Resultatet skal være en dictionary"

    assert result.get(
        "filename"
    ), "filename mangler"

    assert isinstance(
        result.get("file_bytes"),
        bytes,
    ), "file_bytes skal være bytes"

    assert len(
        result["file_bytes"]
    ) > 0, "Excel-filen er tom"

    print(
        "Filnavn:",
        result["filename"],
    )

    print(
        "Antal bytes:",
        len(result["file_bytes"]),
    )

    return result


# -------------------------------------------------
# TEST OVERSKRIFTSRÆKKE
# -------------------------------------------------

def test_find_header(file_result):
    """
    Tester at den detaljerede overskriftsrække findes.
    """

    print(
        "\n2. Læser rå Excel-rækker..."
    )

    raw_rows = read_raw_rule_rows(
        file_result
    )

    assert raw_rows, (
        "Excel-filen gav ingen rækker"
    )

    print(
        "Antal rå rækker:",
        len(raw_rows),
    )

    print(
        "\n3. Finder overskriftsrækken..."
    )

    header_index, headers = (
        find_header_row(raw_rows)
    )

    print(
        "Overskriftsrække fundet ved indeks:",
        header_index,
    )

    print(
        "Kolonnenavne:"
    )

    pprint(headers)

    normalized_headers = {
        str(header).strip().casefold()
        for header in headers
    }

    assert "regelnr" in normalized_headers
    assert "status" in normalized_headers
    assert "regeltype" in normalized_headers
    assert "point0" in normalized_headers
    assert "point5" in normalized_headers

    return raw_rows


# -------------------------------------------------
# TEST NORMALISEREDE REGLER
# -------------------------------------------------

def test_get_mail_rules():
    """
    Tester de færdige MailRule-objekter.
    """

    print(
        "\n4. Henter og normaliserer regler..."
    )

    rules = get_mail_rules()

    assert rules, (
        "Der blev ikke fundet regler"
    )

    print(
        "Antal aktive og testregler:",
        len(rules),
    )

    print(
        "\nFørste fem regler:"
    )

    for rule in rules[:5]:
        pprint(
            rule_to_dict(rule)
        )

    for rule in rules:
        assert isinstance(
            rule.rule_number,
            int,
        )

        assert rule.rule_number > 0

        assert rule.status

        assert rule.status.casefold() != (
            "inaktiv"
        )

    return rules


# -------------------------------------------------
# TEST KENDTE REGLER FRA ARKET
# -------------------------------------------------

# -------------------------------------------------
# TEST KENDTE REGLER FRA ARKET
# -------------------------------------------------

# -------------------------------------------------
# TEST REGELFELTER OG DATATYPER
# -------------------------------------------------

def test_rule_fields_and_types(rules):
    """
    Kontrollerer at reglerne har de rigtige felter
    og datatyper.

    Testen forventer ikke bestemte pointtal.
    Excel-indholdet må gerne ændres over tid.
    """

    print(
        "\n5. Kontrollerer regelfelter og datatyper..."
    )

    valid_statuses = {
        "aktiv",
        "inaktiv",
        "test",
    }

    valid_rule_types = {
        "hovedregel",
        "tillægsregel",
    }

    for rule in rules:
        # Regelnummer skal være et positivt heltal.
        assert isinstance(
            rule.rule_number,
            int,
        ), (
            f"Regelnr skal være int. "
            f"Række={rule.excel_row_number}, "
            f"værdi={rule.rule_number!r}"
        )

        assert rule.rule_number > 0, (
            f"Regelnr skal være større end 0. "
            f"Række={rule.excel_row_number}, "
            f"værdi={rule.rule_number!r}"
        )

        # Status skal være en kendt tekstværdi.
        assert isinstance(
            rule.status,
            str,
        ), (
            f"Status skal være tekst. "
            f"Regelnr={rule.rule_number}, "
            f"værdi={rule.status!r}"
        )

        assert (
            rule.status.casefold()
            in valid_statuses
        ), (
            f"Ukendt status. "
            f"Regelnr={rule.rule_number}, "
            f"værdi={rule.status!r}"
        )

        # Regeltype skal være kendt.
        assert isinstance(
            rule.rule_type,
            str,
        ), (
            f"Regeltype skal være tekst. "
            f"Regelnr={rule.rule_number}, "
            f"værdi={rule.rule_type!r}"
        )

        assert (
            rule.rule_type.casefold()
            in valid_rule_types
        ), (
            f"Ukendt regeltype. "
            f"Regelnr={rule.rule_number}, "
            f"værdi={rule.rule_type!r}"
        )

        # Alle tekstfelter skal være strings.
        text_fields = {
            "destination": rule.destination,
            "subject_pattern": (
                rule.subject_pattern
            ),
            "message_pattern": (
                rule.message_pattern
            ),
            "file_name_pattern": (
                rule.file_name_pattern
            ),
            "file_content_pattern": (
                rule.file_content_pattern
            ),
            "sender_pattern": (
                rule.sender_pattern
            ),
        }

        for field_name, field_value in (
            text_fields.items()
        ):
            assert isinstance(
                field_value,
                str,
            ), (
                f"{field_name} skal være tekst. "
                f"Regelnr={rule.rule_number}, "
                f"værdi={field_value!r}"
            )

        # Alle pointfelter skal være heltal.
        # Blanke Excel-celler skal allerede være 0.
        point_fields = {
            "subject_points": (
                rule.subject_points
            ),
            "message_points": (
                rule.message_points
            ),
            "cpr_points": rule.cpr_points,
            "file_name_points": (
                rule.file_name_points
            ),
            "file_content_points": (
                rule.file_content_points
            ),
            "sender_points": (
                rule.sender_points
            ),
            "max_points": rule.max_points,
        }

        for field_name, field_value in (
            point_fields.items()
        ):
            assert isinstance(
                field_value,
                int,
            ), (
                f"{field_name} skal være int. "
                f"Regelnr={rule.rule_number}, "
                f"værdi={field_value!r}"
            )

    print(
        "Alle regelfelter har korrekte datatyper."
    )

# -------------------------------------------------
# KØR ALLE TESTS
# -------------------------------------------------

def run_all_tests():
    """
    Kører alle tests i rigtig rækkefølge.
    """

    file_result = test_download_rules_file()

    test_find_header(
        file_result
    )

    rules = test_get_mail_rules()

    test_rule_fields_and_types(
        rules
    )

    print(
        "\nAlle tests af mailregler er OK."
    )


# -------------------------------------------------
# KØR TEST DIREKTE
# -------------------------------------------------

if __name__ == "__main__":
    run_all_tests()