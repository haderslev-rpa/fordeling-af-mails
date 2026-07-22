"""
SIMULERET MAILHANDLING

Denne version videresender, flytter eller kategoriserer
ikke mails.

Funktionen returnerer kun, hvad processen senere
skal udføre.

De rigtige Outlook-kald er kommenteret ud.
"""


# -------------------------------------------------
# ER DESTINATION EN MAILADRESSE?
# -------------------------------------------------

def destination_is_email(destination):
    """
    Kontrollerer om destination ligner en mailadresse.
    """

    destination_text = str(
        destination or ""
    ).strip()

    return bool(
        destination_text
        and "@" in destination_text
    )


# -------------------------------------------------
# BYG POINTFORKLARING
# -------------------------------------------------

def build_rule_explanation(result):
    """
    Samler strukturerede point til læsbar tekst.

    Teksten gemmes ikke i item.data.
    """

    return "\n".join(
        (
            (
                f"Emne "
                f"{result['subject_points']} point: "
                f"{result['subject_match']}"
            ),
            (
                f"Besked "
                f"{result['message_points']} point: "
                f"{result['message_match']}"
            ),
            (
                f"Afsender "
                f"{result['sender_points']} point: "
                f"{result['sender_match']}"
            ),
            (
                f"Fil navn "
                f"{result['file_name_points']} point: "
                f"{result['file_name_match']}"
            ),
            (
                f"Fil indhold "
                f"{result['file_content_points']} point: "
                f"{result['file_content_match']}"
            ),
            (
                f"CPR fundet "
                f"{result['cpr_points']} point:"
            ),
            (
                f"Regelnr: "
                f"{result['rule_number']}"
            ),
        )
    )


# -------------------------------------------------
# SIMULÉR HANDLING
# -------------------------------------------------

def execute_mail_action(
    mail,
    rule_result,
):
    """
    Returnerer den handling, som senere skal udføres.

    Denne version ændrer ikke Outlook.
    """

    if not rule_result.get("qualifies"):
        return {
            "action": "no_match",
            "destination": None,
            "status_message": (
                "Ingen regel matchede mailen"
            ),
            "simulated": True,
        }

    destination = (
        rule_result.get("destination")
        or ""
    )

    rule_status = (
        rule_result.get("rule_status")
        or ""
    ).casefold()

    # -------------------------------------------------
    # TESTREGEL
    # -------------------------------------------------

    if rule_status == "test":
        category = (
            f"Test - Nr: "
            f"{rule_result['rule_number']} - "
            f"Point: "
            f"{rule_result['total_points']}"
        )

        # Aktiveres senere:
        #
        # from q_outlook_api.functionality.mail_api import (
        #     add_mail_category,
        # )
        #
        # add_mail_category(
        #     user_mail=mail["mailbox"],
        #     message_id=mail["message_id"],
        #     category=category,
        # )

        return {
            "action": "test_category",
            "destination": category,
            "status_message": (
                "Testkategori ville blive tilføjet"
            ),
            "simulated": True,
        }

    # -------------------------------------------------
    # VIDERESENDELSE
    # -------------------------------------------------

    if destination_is_email(
        destination
    ):
        explanation = build_rule_explanation(
            rule_result
        )

        # Aktiveres senere:
        #
        # from q_outlook_api.functionality.mail_api import (
        #     forward_mail,
        # )
        #
        # forward_mail(
        #     user_mail=mail["mailbox"],
        #     message_id=mail["message_id"],
        #     forward={
        #         "to": [destination],
        #         "cc": [],
        #         "bcc": [],
        #         "body": explanation,
        #     },
        # )

        return {
            "action": "forward",
            "destination": destination,
            "status_message": (
                "Mail ville blive videresendt"
            ),
            "explanation": explanation,
            "simulated": True,
        }

    # -------------------------------------------------
    # FLYT TIL UNDERMAPPE
    # -------------------------------------------------

    if destination:
        # Aktiveres senere:
        #
        # folder_id = find_folder_id(...)
        #
        # move_mail(
        #     user_mail=mail["mailbox"],
        #     message_id=mail["message_id"],
        #     destination_folder_id=folder_id,
        # )

        return {
            "action": "move",
            "destination": destination,
            "status_message": (
                "Mail ville blive flyttet "
                "til undermappe"
            ),
            "simulated": True,
        }

    return {
        "action": "no_action",
        "destination": None,
        "status_message": (
            "Regel matchede uden destination"
        ),
        "simulated": True,
    }