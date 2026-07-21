import asyncio
import logging
import sys
from pprint import pprint

from behandel import behandel_page

from automation_server_client import (
    AutomationServer,
    WorkItemError,
    WorkItemStatus,
    Workqueue,
)

from q_haderslev_vbo.automation_server.ats_update_item_data import (
    update_item_data,
)

from hent_mailregler_fra_sharepoint import (
    get_mail_rules,
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

        workqueue.add_item(
            data=data_json,
            reference=(
                data_json["box"]["mail"]["received_datetime_danish"]
                + " - "
                + data_json["box"]["mail"]["mailbox"]
                + " - "
                + data_json["box"]["mail"]["subject"][:50]
            ),
        )

        if debug:
            logger.info(
                "Mail tilføjet til queue: %s",
                data_json["box"]["mail"]["subject"],
            )

    logger.info(
        "%s mails tilføjet til workqueue",
        len(box_data_items),
    )
# ------------------------------------------------------------
# PROCESS-MODE
# ------------------------------------------------------------

async def process_workqueue(
    workqueue: Workqueue,
    debug: bool,
):
    """
    Henter regler én gang og behandler items.
    """

    logger = logging.getLogger(__name__)

    logger.info(
        "Process workqueue mode started "
        "(debug=%s)",
        debug,
    )

    rules = get_mail_rules()

    logger.info(
        "%s regler blev hentet fra Excel",
        len(rules),
    )

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

                await behandel_page(
                    item=item,
                    rules=rules,
                    debug=debug,
                )

                update_item_data(
                    data,
                    item=item,
                    status="Completed",
                    status_code="Færdig",
                )

                item.update(data)
                item.complete("Completed")

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
    QUEUE_MODE = "--queue" in sys.argv

    ats = AutomationServer.from_environment()
    workqueue = ats.workqueue()

    if QUEUE_MODE:
        # VIGTIGT:
        # Denne linje sletter alle NEW items.
        #
        # Kommentér linjen ud, hvis eksisterende
        # NEW items skal bevares.
        workqueue.clear_workqueue(
            WorkItemStatus.NEW
        )

        asyncio.run(
            populate_queue(
                workqueue,
                debug=DEBUG,
            )
        )

        sys.exit(0)

    asyncio.run(
        process_workqueue(
            workqueue,
            debug=DEBUG,
        )
    )