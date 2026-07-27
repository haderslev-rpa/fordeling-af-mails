"""
HENT MAILS TIL QUEUE

Denne fil:
- henter mails fra Outlook
- filtrerer robotbehandlede mails
- bygger box_data
- returnerer box_data til main.py

Denne fil:
- tilføjer ikke items til Automation Server
- laver ikke queue-reference
- bruger ikke workqueue.add_item
- bruger ikke update_item_data
- henter ikke Excel-regler
"""

import logging

from q_outlook_api.functionality.mail_api import (
    get_mails,
)

from proces_konfiguration import (
    MAILBOXES,
    MAIL_LIMIT_PER_MAILBOX,
    ROBOT_CATEGORY_MARKERS,
    ROBOT_RULE_NUMBER_MARKER,
    build_previous_forward_text,
)


logger = logging.getLogger(__name__)


# -------------------------------------------------
# ROBOTKATEGORIER
# -------------------------------------------------

def har_robotkategori(categories):
    """
    Kontrollerer om mailen har en robotkategori.

    categories er en liste (samling af værdier).
    """

    for category in categories or []:
        category_text = str(category)

        for marker in ROBOT_CATEGORY_MARKERS:
            if marker in category_text:
                return True

    return False


# -------------------------------------------------
# TIDLIGERE ROBOTVIDERESENDELSE
# -------------------------------------------------

def er_tidligere_videresendt_af_robot(
    body,
    mailbox,
):
    """
    Kontrollerer de samme to betingelser
    som den eksisterende Blue Prism-proces.

    Begge tekster skal findes:
    - den faste videresendelsestekst
    - teksten Regelnr:
    """

    body_text = str(body or "")

    expected_text = build_previous_forward_text(
        mailbox
    )

    return (
        expected_text in body_text
        and ROBOT_RULE_NUMBER_MARKER in body_text
    )


# -------------------------------------------------
# BYG BOX_DATA
# -------------------------------------------------

def byg_box_data(
    mail,
    mailbox,
):
    """
    Bygger værdierne, som main.py placerer
    under item.data["box"].

    Mail-body, kategorier og filindhold
    bliver ikke gemt.
    """

    attachment_names = list(
        mail.get("attachment_names")
        or []
    )

    return {
        "mail": {
            "mailbox": mailbox,

            "message_id": mail.get(
                "message_id"
            ),

            "internet_message_id": mail.get(
                "internet_message_id"
            ),

            "subject": (
                mail.get("subject")
                or ""
            ),

            "sender_address": mail.get(
                "sender_address"
            ),

            "received_datetime_utc": mail.get(
                "received_datetime_utc"
            ),

            "received_datetime_danish": mail.get(
                "received_datetime_danish"
            ),

            "has_attachments": (
                len(attachment_names) > 0
            ),

            "attachment_count": len(
                attachment_names
            ),

            "attachment_names": (
                attachment_names
            ),
        },

        # Den vindende regel gemmes senere.
        "subject_points": None,
        "subject_match": "",

        "message_points": None,
        "message_match": "",

        "sender_points": None,
        "sender_match": "",

        "file_name_points": None,
        "file_name_match": "",

        "file_content_points": None,
        "file_content_match": "",

        "cpr_points": None,
        "cpr_found": None,

        "rule_number": None,
        "total_points": None,

        "action": None,
        "destination": None,

        "azure_attempt_count": 0,
    }


# -------------------------------------------------
# HENT MAILS FRA ÉN POSTKASSE
# -------------------------------------------------

def hent_mails_fra_postkasse(
    mailbox_config,
):
    """
    Henter mails fra én postkasse.

    Body hentes som almindelig tekst, fordi body
    bruges midlertidigt til Blue Prism-filteret.

    Body gemmes ikke i queue-itemet.
    """
    logger.info(
        "Henter mails fra den faktiske postkasse: %s",
        mailbox_config.address,
)

    return get_mails(
        user_mail=mailbox_config.address,
        folder=mailbox_config.folder,
        limit=MAIL_LIMIT_PER_MAILBOX,
        include_attachments=True,
        get_inline=False,
        prefer_plain_text=True,
    )


# -------------------------------------------------
# HENT MAILS TIL QUEUE
# -------------------------------------------------

def hent_mails_til_queue():
    """
    Henter og filtrerer mails fra alle aktive
    postkasser.

    Funktionen returnerer en liste (samling)
    med box_data.

    main.py sørger selv for:
    - update_item_data
    - queue-reference
    - workqueue.add_item
    """

    box_data_items = []

    skipped_category_count = 0
    skipped_forward_count = 0
    skipped_missing_id_count = 0

    for mailbox_config in MAILBOXES:
        if not mailbox_config.enabled:
            continue

        logger.info(
            "Henter mails fra %s",
            mailbox_config.address,
        )

        mails = hent_mails_fra_postkasse(
            mailbox_config
        )

        logger.info(
            "%s mails blev hentet fra %s",
            len(mails),
            mailbox_config.address,
        )

        for mail in mails:
            message_id = mail.get(
                "message_id"
            )

            if not message_id:
                skipped_missing_id_count += 1

                logger.warning(
                    "Mail uden message_id blev "
                    "sprunget over. Emne: %s",
                    mail.get("subject"),
                )

                continue

            if har_robotkategori(
                mail.get("categories")
            ):
                skipped_category_count += 1

                logger.debug(
                    "Mail sprunget over på grund "
                    "af robotkategori: %s",
                    mail.get("subject"),
                )

                continue

            if er_tidligere_videresendt_af_robot(
                body=mail.get("body"),
                mailbox=mailbox_config.address,
            ):
                skipped_forward_count += 1

                logger.debug(
                    "Tidligere robotvideresendt mail "
                    "sprunget over: %s",
                    mail.get("subject"),
                )

                continue

            box_data = byg_box_data(
                mail=mail,
                mailbox=mailbox_config.address,
            )

            box_data_items.append(
                box_data
            )

    # -------------------------------------------------
    # SORTÉR ÆLDSTE MAIL FØRST
    # -------------------------------------------------

    box_data_items.sort(
        key=lambda item: (
            item["mail"].get(
                "received_datetime_utc"
            )
            or ""
        )
    )

    logger.info(
        (
            "Mailhentning færdig. "
            "Klargjort til queue: %s. "
            "Robotkategori: %s. "
            "Tidligere videresendt: %s. "
            "Mangler message_id: %s."
        ),
        len(box_data_items),
        skipped_category_count,
        skipped_forward_count,
        skipped_missing_id_count,
    )

    return box_data_items