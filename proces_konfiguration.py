"""
CENTRAL KONFIGURATION

Denne fil indeholder faste indstillinger til processen.

Credentials må ikke placeres i denne fil.
"""

import os
from dataclasses import dataclass


# -------------------------------------------------
# POSTKASSEKONFIGURATION
# -------------------------------------------------

@dataclass(frozen=True)
class MailboxConfig:
    """
    Konfiguration for én postkasse.
    """

    address: str
    folder: str = "inbox"
    enabled: bool = True


MAILBOXES = [
    MailboxConfig(
        address="jobcenter@haderslev.dk",
        folder="inbox",
        enabled=True,
    ),

    # Aktiveres senere.
    MailboxConfig(
        address="post@haderslev.dk",
        folder="inbox",
        enabled=False,
    ),
]


# -------------------------------------------------
# QUEUE
# -------------------------------------------------

QUEUE_NAME = os.getenv(
    "MAIL_QUEUE_NAME",
    "Fordeling af mails",
)

# None betyder, at Graph-pagination fortsætter,
# indtil alle mails er hentet.
MAIL_LIMIT_PER_MAILBOX = None


# -------------------------------------------------
# SHAREPOINT OG EXCEL
# -------------------------------------------------

RULES_SITE_NAME = "Automatisering"

RULES_FILE_PATH = (
    "RPA - Processer/"
    "Fordeling af mails/"
    "Fællespostkasse - regler til emails.xlsx"
)

RULES_SHEET_NAME = "Regler"


# -------------------------------------------------
# ROBOTKATEGORIER
# -------------------------------------------------

ROBOT_CATEGORY_MARKERS = (
    "Point: ",
    "Test",
    "Nr:",
)


# -------------------------------------------------
# TIDLIGERE ROBOTVIDERESENDELSE
# -------------------------------------------------

ROBOT_RULE_NUMBER_MARKER = "Regelnr:"


def build_previous_forward_text(mailbox):
    """
    Bygger Blue Prism-teksten til kontrol af
    tidligere robotvideresendelse.
    """

    return (
        f"Videresend altid til {mailbox} "
        f"<mailto:{mailbox}>  . "
        "Skriv gerne hvem du tror der skulle "
        "have modtaget mailen og hvorfor."
    )


# -------------------------------------------------
# REGELSTATUS
# -------------------------------------------------

RULE_STATUS_ACTIVE = "Aktiv"
RULE_STATUS_INACTIVE = "Inaktiv"
RULE_STATUS_TEST = "Test"


# -------------------------------------------------
# REGELTYPE
# -------------------------------------------------

RULE_TYPE_MAIN = "Hovedregel"
RULE_TYPE_ADDITIONAL = "Tillægsregel"


# -------------------------------------------------
# POINT
# -------------------------------------------------

RULE_MINIMUM_POINTS = 300


# -------------------------------------------------
# AZURE
# -------------------------------------------------

MAX_AZURE_ATTEMPTS_PER_ITEM = 50


# -------------------------------------------------
# STATES TIL SENERE WORKER-BEHANDLING
# -------------------------------------------------

STATE_MAIL_FETCHED = "1.0 - Mail hentet"

STATE_ATTACHMENTS_FETCHED = (
    "2.0 - Vedhæftninger hentet"
)

STATE_FILE_CONTENT_READ = (
    "3.0 - Filindhold læst"
)

STATE_AZURE_LIMIT_REACHED = (
    "3.2 - Maksimalt antal Azure-forsøg nået"
)

STATE_RULES_EVALUATED = (
    "4.0 - Regler vurderet"
)

STATE_MAIL_PROCESSED = (
    "5.0 - Mail behandlet"
)