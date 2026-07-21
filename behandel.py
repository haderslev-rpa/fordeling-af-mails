"""
FØRSTE WORKER-TEST

Denne version bruges kun til at kontrollere:

1. Worker kan hente et queue-item.
2. box.mail findes.
3. Reglerne er hentet fra Excel.
4. Item-data kan opdateres gennem update_item_data.

Der hentes endnu ingen attachments.
Der udføres endnu ingen mailhandling.
"""

import logging

from automation_server_client import (
    WorkItemError,
)

from q_haderslev_vbo.automation_server.ats_update_item_data import (
    update_item_data,
)


logger = logging.getLogger(__name__)


# -------------------------------------------------
# KONTROLLÉR ITEM.DATA
# -------------------------------------------------

def validate_item_data(item):
    """
    Kontrollerer den forventede item-struktur.
    """

    data = item.data or {}

    box = data.get("box")

    if not isinstance(box, dict):
        raise WorkItemError(
            "Item mangler box."
        )

    mail = box.get("mail")

    if not isinstance(mail, dict):
        raise WorkItemError(
            "Item mangler box.mail."
        )

    if not mail.get("mailbox"):
        raise WorkItemError(
            "Item mangler box.mail.mailbox."
        )

    if not mail.get("message_id"):
        raise WorkItemError(
            "Item mangler box.mail.message_id."
        )

    return data, mail


# -------------------------------------------------
# BEHANDEL ÉT ITEM
# -------------------------------------------------

async def behandel_page(
    item,
    rules,
    debug=False,
):
    """
    Første worker-test.

    Funktionen kontrollerer queue-item'et
    og antallet af indlæste regler.
    """

    data, mail = validate_item_data(
        item
    )

    logger.info(
        "Worker har hentet mail-reference: %s",
        mail.get("subject"),
    )

    logger.info(
        "%s regler er tilgængelige i worker",
        len(rules),
    )

    if debug:
        logger.info(
            "Postkasse: %s",
            mail.get("mailbox"),
        )

        logger.info(
            "Message ID: %s",
            mail.get("message_id"),
        )

        if rules:
            logger.info(
                "Første regelnummer: %s",
                rules[0].rule_number,
            )

    # Vi opdaterer ikke state eller status endnu.
    # Det sker først, når den egentlige behandling
    # er implementeret.

    update_item_data(
        data,
        box_updates={
            "loaded_rule_count": len(
                rules
            ),
        },
        item=item,
    )

    return {
        "mailbox": mail.get("mailbox"),
        "message_id": mail.get(
            "message_id"
        ),
        "loaded_rule_count": len(
            rules
        ),
    }