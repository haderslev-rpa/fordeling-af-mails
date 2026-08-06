import asyncio
import logging
import sys
from pprint import pprint
import re
import unicodedata



from automation_server_client import (
    AutomationServer,
    WorkItemError,
    WorkItemStatus,
    Workqueue,
)

from q_haderslev_vbo.automation_server.ats_update_item_data import (
    update_item_data,
)
from q_haderslev_vbo.automation_server.ats_is_item_in_queue import (
    is_item_in_queue,
)


from hent_mails_til_queue import (
    hent_mails_til_queue,
)




# ------------------------------------------------------------
# LOGGING
# ------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format=(
        "%(asctime)s "
        "[%(levelname)s] "
        "%(name)s: "
        "%(message)s"
    ),
)

logging.getLogger(
    "httpx"
).setLevel(logging.WARNING)

logging.getLogger(
    "automation_server_client"
).setLevel(logging.WARNING)

logging.getLogger(
    "debugpy"
).setLevel(logging.WARNING)


# ------------------------------------------------------------
# NORMALISÉR REFERENCETEKST
# ------------------------------------------------------------

def normalize_reference_text(
    value,
):
    """
    Gør tekst stabil til queue-reference.

    Bevarer:
    - æ
    - ø
    - å

    Fjerner forskelle i:
    - tabs
    - linjeskift
    - dobbelt mellemrum
    - mærkelige unicode-varianter
    """

    text = str(
        value or ""
    )

    text = unicodedata.normalize(
        "NFKC",
        text,
    )

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text.strip()


# ------------------------------------------------------------
# BUILD ITEM REFERENCE
# ------------------------------------------------------------

def build_item_reference(
    mail_data,
):
    """
    Bygger stabil queue-reference.
    """

    return (
        normalize_reference_text(
            mail_data[
                "received_datetime_danish"
            ]
        )
        + " - "
        + normalize_reference_text(
            mail_data[
                "mailbox"
            ]
        )
        + " - "
        + normalize_reference_text(
            mail_data[
                "subject"
            ]
        )[:50]
    )


# ------------------------------------------------------------
# QUEUE-MODE (PRODUCER)
# ------------------------------------------------------------

async def populate_queue(
    workqueue: Workqueue,
    debug: bool,
):
    """
    Henter mails og tilføjer dem til workqueue.

    """

    logger = logging.getLogger(__name__)

    logger.info(
        "Populate queue mode started"
    )

    box_data_items = hent_mails_til_queue()

    for box_data in box_data_items:
        data_json = {}

        update_item_data(
            data_json,
            box_updates=box_data,
            update=False,
        )

        item_reference = (
            build_item_reference(
                data_json["box"]["mail"]
            )
        )

        # ------------------------------------------------------------
        # DUBLETKONTROL
        # ------------------------------------------------------------

        force_reprocess = bool(
            data_json["box"].get(
                "force_reprocess",
                False,
            )
        )

        if force_reprocess:
            print(
                "GENBEHANDLER ALTID:",
                item_reference,
            )

        else:
            exists = is_item_in_queue(
                queue_id=workqueue.id,
                item_reference=item_reference,

                new=True,
                in_progress=True,
                completed=True,
                failed=False,
                pending_user_action=True,

                updated_at=False,
            )

            if exists:
                print(
                    "SPRINGER OVER:",
                    item_reference,
                )

                continue

        workqueue.add_item(
            data=data_json,
            reference=item_reference,
        )

        logger.info(
            "Item tilføjet: %s",
            item_reference,
        )

# ------------------------------------------------------------
# PROCESS-MODE (WORKER)
# ------------------------------------------------------------

async def process_workqueue(
    workqueue: Workqueue,
    debug: bool,
):
    """
    Henter alle regler én gang og behandler items.

    behandel.py:
        Udfører procestrinnene.
        Opdaterer box og states.
        Returnerer slutbeskeden.

    main.py:
        Opdaterer status og status_code.
        Afslutter item'et.
    """

    logger = logging.getLogger(__name__)

    logger.info(
        "Process workqueue mode started "
        "(debug=%s)",
        debug,
    )

    # --------------------------------------------------------
    # WORKER-IMPORTS
    #
    # Importeres kun i worker-mode.
    # --queue indlæser derfor ikke worker-filerne.
    # --------------------------------------------------------

    from behandel import behandel_page

    from hent_mailregler_fra_sharepoint import (
        get_all_mail_rules,
    )

    # --------------------------------------------------------
    # HENT ALLE REGLER ÉN GANG
    # --------------------------------------------------------

    all_rules = get_all_mail_rules()

    logger.info(
        "%s regler blev hentet fra Excel",
        len(all_rules),
    )

    # --------------------------------------------------------
    # BEHANDL ITEMS
    # --------------------------------------------------------

    for item in workqueue:
        with item:
            data = item.data

            try:
                print(
                    "==================================== "
                    "NEXT ITEM "
                    "===================================="
                )

                pprint(data)

                process_result = await behandel_page(
                    item=item,
                    all_rules=all_rules,
                    debug=debug,
                )

                # Slutbeskeden kommer fra behandel.py.
                # Hvis resultatet mod forventning er tomt,
                # bruges en neutral standardbesked.
                status_message = (
                    (
                        process_result
                        or {}
                    ).get(
                        "status_message"
                    )
                    or "Mail behandlet"
                )

                # Status og status_code sættes samlet,
                # præcis som update_item_data kræver.
                update_item_data(
                    data,
                    item=item,
                    status=status_message,
                    status_code="Færdig",
                )

                item.update(data)

                comparison_message = (
                    process_result.get(
                        "comparison_message"
                    )
                    or "Completed"
                )

                item.complete(
                    comparison_message
                )

            except WorkItemError as error:
                logger.error(
                    "WorkItemError for item %s: %s",
                    item.reference,
                    error,
                )

                item.fail(
                    str(error)
                )

            except Exception:
                logger.exception(
                    "Uventet fejl"
                )

                raise

# ------------------------------------------------------------
# MAIN ENTRY POINT
# ------------------------------------------------------------

if __name__ == "__main__":

    DEBUG = "--debug" in sys.argv

    QUEUE_MODE = (
        "--queue" in sys.argv
    )

    QUEUE_AND_PROCESS_MODE = (
        "--queue-and-process"
        in sys.argv
    )

    ats = AutomationServer.from_environment()

    workqueue = ats.workqueue()

    if QUEUE_MODE or QUEUE_AND_PROCESS_MODE:

        # VIGTIGT:
        # Denne linje sletter alle NEW items.
        #
        # Kommentér linjen ud, hvis eksisterende
        # NEW items skal bevares.
        #
        # workqueue.clear_workqueue(
        #     WorkItemStatus.NEW
        # )

        asyncio.run(
            populate_queue(
                workqueue,
                debug=DEBUG,
            )
        )

        # Stop kun hvis det er ren queue-mode
        if QUEUE_MODE:

            sys.exit(0)

    asyncio.run(
        process_workqueue(
            workqueue,
            debug=DEBUG,
        )
    )