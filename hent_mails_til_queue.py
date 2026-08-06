"""
HENT MAILS TIL QUEUE

Denne fil:
- henter mails fra Outlook
- filtrerer robotbehandlede mails
- finder manuel genbehandling
- bygger box_data
- sorterer mails med ældste først
- returnerer box_data til main.py

Denne fil:
- tilføjer ikke items til Automation Server
- laver ikke queue-reference
- bruger ikke workqueue.add_item
- bruger ikke update_item_data
- henter ikke Excel-regler

MANUEL GENBEHANDLING:

Hvis brugeren sætter kategorien:

Robot genbehandel

så må mailen gå videre, selv om:
- mailen har andre robotkategorier
- mailens body indeholder tidligere robottekst
- mailen tidligere er behandlet

Kategorien fjernes ikke i denne fil.

Når worker-flowet behandler mailen med kategorier,
overskrives alle kategorier med det nye resultat.
"""

import html
import logging
import re
import unicodedata

from q_outlook_api.functionality.mail_api import (
    get_mails,
)

from proces_konfiguration import (
    MAILBOXES,
    MAIL_LIMIT_PER_MAILBOX,
    ROBOT_CATEGORY_MARKERS,
    ROBOT_REPROCESS_CATEGORY,
    ROBOT_RULE_NUMBER_MARKER,
    build_previous_forward_text,
)


logger = logging.getLogger(__name__)


# -------------------------------------------------
# MANUEL GENBEHANDLING
# -------------------------------------------------

def skal_genbehandles(
    categories,
):
    """
    Kontrollerer om brugeren har valgt
    kategorien "Robot genbehandel".

    Sammenligningen ignorerer:
    - store og små bogstaver
    - mellemrum før kategorien
    - mellemrum efter kategorien
    """

    expected_category = (
        ROBOT_REPROCESS_CATEGORY
        .strip()
        .casefold()
    )

    return any(
        str(category)
        .strip()
        .casefold()
        == expected_category
        for category in (
            categories
            or []
        )
    )


# -------------------------------------------------
# ROBOTKATEGORIER
# -------------------------------------------------

def har_robotkategori(
    categories,
):
    """
    Kontrollerer om mailen har en kategori,
    som robotten tidligere har skrevet.

    Følgende kategorier genkendes:

    Test
    Point: 0
    Point: 150
    Nr: 3
    Nr: 1035

    Sammenligningen ignorerer forskel på
    store og små bogstaver samt mellemrum.
    """

    for category in categories or []:
        category_text = (
            str(category)
            .strip()
            .casefold()
        )

        if category_text == "test":
            return True

        if category_text.startswith(
            "point:"
        ):
            return True

        if category_text.startswith(
            "nr:"
        ):
            return True

    return False

# -------------------------------------------------
# NORMALISÉR BODY TIL KONTROL
# -------------------------------------------------

def normaliser_body_til_kontrol(
    body,
):
    """
    Normaliserer mailens body til kontrol.

    Funktionen håndterer:
    - HTML-koder
    - Unicode-varianter
    - linjeskift
    - tabs
    - dobbelte mellemrum
    - store og små bogstaver
    """

    body_text = html.unescape(
        str(
            body
            or ""
        )
    )

    body_text = unicodedata.normalize(
        "NFKC",
        body_text,
    )

    body_text = re.sub(
        r"\s+",
        " ",
        body_text,
    )

    return (
        body_text
        .strip()
        .casefold()
    )


# -------------------------------------------------
# TIDLIGERE ROBOTVIDERESENDELSE
# -------------------------------------------------

def er_tidligere_videresendt_af_robot(
    body,
    mailbox,
):
    """
    Kontrollerer samme princip som Blue Prism.

    Mailen betragtes kun som tidligere behandlet,
    når body både indeholder:

    1. "Videresend altid til [aktuel postkasse]"
    2. "Regelnr:"

    Det betyder:

    - En mail, som vender tilbage til samme
      postkasse, bliver sprunget over.

    - En mail, som videresendes fra én overvåget
      postkasse til en anden, må behandles i den
      nye postkasse.

    Manuel kategori "Robot genbehandel" håndteres
    før denne funktion og tilsidesætter kontrollen.
    """

    body_text = normaliser_body_til_kontrol(
        body
    )

    mailbox_text = (
        str(
            mailbox
            or ""
        )
        .strip()
        .casefold()
    )

    if not mailbox_text:
        return False


    # -------------------------------------------------
    # KONTROLLÉR REGELNUMMER
    # -------------------------------------------------

    rule_number_found = (
        ROBOT_RULE_NUMBER_MARKER
        .strip()
        .casefold()
        in body_text
    )


    # -------------------------------------------------
    # KONTROLLÉR AKTUEL POSTKASSE
    # -------------------------------------------------

    current_mailbox_text_found = (
        (
            "videresend altid til "
            f"{mailbox_text}"
        )
        in body_text
    )


    # -------------------------------------------------
    # RESULTAT
    # -------------------------------------------------

    return (
        current_mailbox_text_found
        and rule_number_found
    )
# -------------------------------------------------
# BYG BOX_DATA
# -------------------------------------------------

def byg_box_data(
    mail,
    mailbox,
    force_reprocess=False,
):
    """
    Bygger værdierne, som main.py placerer
    under item.data["box"].

    Mail-body, kategorier og filindhold
    bliver ikke gemt.

    force_reprocess=True betyder, at brugeren
    har valgt kategorien "Robot genbehandel".
    """

    attachment_names = list(
        mail.get(
            "attachment_names"
        )
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

        # Brugeren har manuelt valgt,
        # at mailen skal behandles igen.
        "force_reprocess": bool(
            force_reprocess
        ),

        # Den valgte regel gemmes senere.
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
    bruges midlertidigt til robotfilteret.

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

    Funktionen returnerer en liste med box_data.

    main.py sørger selv for:
    - update_item_data
    - queue-reference
    - dubletkontrol
    - workqueue.add_item

    Der udskrives kun én samlet oversigt,
    når alle postkasser er gennemgået.
    """

    box_data_items = []

    total_mail_count = 0
    skipped_category_count = 0
    skipped_forward_count = 0
    skipped_missing_id_count = 0
    reprocess_count = 0

    for mailbox_config in MAILBOXES:
        if not mailbox_config.enabled:
            continue

        mails = hent_mails_fra_postkasse(
            mailbox_config
        )

        total_mail_count += len(
            mails
        )

        for mail in mails:
            message_id = mail.get(
                "message_id"
            )

            if not message_id:
                skipped_missing_id_count += 1

                continue

            categories = (
                mail.get("categories")
                or []
            )

            force_reprocess = (
                skal_genbehandles(
                    categories
                )
            )

            # -------------------------------------------------
            # MANUEL GENBEHANDLING
            # -------------------------------------------------

            if force_reprocess:
                reprocess_count += 1

            # -------------------------------------------------
            # ROBOTKATEGORI
            # -------------------------------------------------

            elif har_robotkategori(
                categories
            ):
                skipped_category_count += 1

                continue

            # -------------------------------------------------
            # TIDLIGERE ROBOTVIDERESENDELSE
            # -------------------------------------------------

            if (
                not force_reprocess
                and er_tidligere_videresendt_af_robot(
                    body=mail.get("body"),
                    mailbox=(
                        mailbox_config.address
                    ),
                )
            ):
                skipped_forward_count += 1

                continue

            # -------------------------------------------------
            # BYG BOX_DATA
            # -------------------------------------------------

            box_data = byg_box_data(
                mail=mail,
                mailbox=(
                    mailbox_config.address
                ),
                force_reprocess=(
                    force_reprocess
                ),
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


    # -------------------------------------------------
    # RESULTAT
    # -------------------------------------------------

    print()
    print("===================================")
    print("RESULTAT AF MAILFILTRERING")
    print("===================================")
    print(
        "Mails hentet:",
        total_mail_count,
    )
    print(
        "Sendes videre til main:",
        len(box_data_items),
    )
    print(
        "Manuel genbehandling:",
        reprocess_count,
    )
    print(
        "Robotkategori:",
        skipped_category_count,
    )
    print(
        "Tidligere videresendt:",
        skipped_forward_count,
    )
    print(
        "Mangler message_id:",
        skipped_missing_id_count,
    )
    print("===================================")
    print()

    return box_data_items
