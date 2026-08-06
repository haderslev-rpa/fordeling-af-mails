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
    Bestemmer hvilken handling processen ville udføre.

    Der foretages ingen rigtig Outlook-handling endnu.

    Mindst 300 point:
        Simulér videresendelse, flytning eller testkategori.

    Under 300 point:
        Simulér kun kategorier.
        Mailen må ikke videresendes eller flyttes.
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

    destination = (
        rule_result.get(
            "destination"
        )
        or ""
    )

    rule_status = (
        rule_result.get(
            "rule_status"
        )
        or ""
    ).casefold()

    # -------------------------------------------------
    # INGEN REGEL ELLER 0 POINT
    # -------------------------------------------------

    if (
        rule_number is None
        or total_points <= 0
    ):
        return {
            "action": "category_only",
            "destination": None,
            "categories": [
                "Point: 0",
            ],
            "status_message": (
                "Ingen regel gav point"
            ),
            "simulated": True,
        }

    # -------------------------------------------------
    # BEDSTE REGEL UNDER 300 POINT
    # -------------------------------------------------

    if total_points < 300:
        categories = [
            f"Point: {total_points}",
            f"Nr: {rule_number}",
        ]

        # Aktiveres senere, når testen er godkendt:
        #
        # from q_outlook_api.functionality.mail_api import (
        #     update_mail_categories,
        # )
        #
        # update_mail_categories(
        #     user_mail=mail["mailbox"],
        #     message_id=mail["message_id"],
        #     categories=categories,
        # )

        return {
            "action": "category_only",
            "destination": None,
            "categories": categories,
            "status_message": (
                "Kategorier ville blive tilføjet. "
                "Mailen ville ikke blive fordelt"
            ),
            "simulated": True,
        }

    # -------------------------------------------------
    # TESTREGEL MED MINDST 300 POINT
    # -------------------------------------------------

    if rule_status == "test":
        categories = [
            "Test",
            f"Point: {total_points}",
            f"Nr: {rule_number}",
        ]

        # Aktiveres senere:
        #
        # from q_outlook_api.functionality.mail_api import (
        #     update_mail_categories,
        # )
        #
        # update_mail_categories(
        #     user_mail=mail["mailbox"],
        #     message_id=mail["message_id"],
        #     categories=categories,
        # )

        return {
            "action": "test_category",
            "destination": None,
            "categories": categories,
            "status_message": (
                "Testkategorier ville blive tilføjet"
            ),
            "simulated": True,
        }

    # -------------------------------------------------
    # VIDERESENDELSE MED MINDST 300 POINT
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
            "categories": [],
            "status_message": (
                "Mail ville blive videresendt"
            ),
            "explanation": explanation,
            "simulated": True,
        }

    # -------------------------------------------------
    # FLYTNING MED MINDST 300 POINT
    # -------------------------------------------------

    if destination:
        # Aktiveres senere:
        #
        # folder_id = find_folder_id(
        #     mailbox=mail["mailbox"],
        #     folder_name=destination,
        # )
        #
        # move_mail(
        #     user_mail=mail["mailbox"],
        #     message_id=mail["message_id"],
        #     destination_folder_id=folder_id,
        # )

        return {
            "action": "move",
            "destination": destination,
            "categories": [],
            "status_message": (
                "Mail ville blive flyttet "
                "til undermappe"
            ),
            "simulated": True,
        }

    # -------------------------------------------------
    # REGEL UDEN DESTINATION
    # -------------------------------------------------

    return {
        "action": "no_action",
        "destination": None,
        "categories": [],
        "status_message": (
            "Regel nåede 300 point, "
            "men manglede destination"
        ),
        "simulated": True,
    }

# -------------------------------------------------
# BUILD COMPARISON MESSAGE
# -------------------------------------------------

def build_comparison_message(
    rule_result,
    action_result,
):
    """
    Bygger beskeden til Automation Server.

    Mindst 300 point:
        Nr: 1035 + byg@haderslev.dk

    Under 300 point:
        Point: 150

    Ingen point:
        Point: 0
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

    destination = action_result.get(
        "destination"
    )

    # -------------------------------------------------
    # UNDER 300 POINT
    # -------------------------------------------------

    if total_points < 300:
        return (
            f"Point: {total_points}"
        )

    # -------------------------------------------------
    # MINDST 300 POINT
    # -------------------------------------------------

    if rule_number is not None:
        message = (
            f"Nr: {rule_number}"
        )

        if destination:
            message += (
                f" + {destination}"
            )

        return message

    return (
        f"Point: {total_points}"
    )
