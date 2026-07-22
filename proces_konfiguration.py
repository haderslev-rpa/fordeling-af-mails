"""
CENTRAL KONFIGURATION

TEST_MAILBOX_OVERRIDE:
    Bestemmer hvilken Outlook-postkasse mails hentes fra.

RULE_MAILBOX_OVERRIDE:
    Bestemmer hvilken postkasses Excel-regler worker bruger.

Eksempel i .env:

TEST_MAILBOX_OVERRIDE=robot-data@haderslev.dk
RULE_MAILBOX_OVERRIDE=jobcenter@haderslev.dk

Det betyder:

Mails hentes fra:
robot-data@haderslev.dk

Regler hentes som om mailen tilhører:
jobcenter@haderslev.dk
"""

import os
from dataclasses import dataclass

from dotenv import load_dotenv


# -------------------------------------------------
# INDLÆS .ENV
# -------------------------------------------------

load_dotenv()


# -------------------------------------------------
# TESTOVERRIDES
# -------------------------------------------------

TEST_MAILBOX_OVERRIDE = os.getenv(
    "TEST_MAILBOX_OVERRIDE",
    "",
).strip()

RULE_MAILBOX_OVERRIDE = os.getenv(
    "RULE_MAILBOX_OVERRIDE",
    "",
).strip()


# -------------------------------------------------
# POSTKASSER
# -------------------------------------------------

@dataclass(frozen=True)
class MailboxConfig:
    """
    Konfiguration for én Outlook-postkasse.
    """

    address: str
    folder: str = "inbox"
    enabled: bool = True


MAILBOXES = [
    MailboxConfig(
        address=(
            TEST_MAILBOX_OVERRIDE
            or "jobcenter@haderslev.dk"
        ),
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


# Brug et lille tal under test.
#
# Eksempel:
# MAIL_LIMIT_PER_MAILBOX = 5
#
# None betyder, at alle mails hentes.
MAIL_LIMIT_PER_MAILBOX = 200


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
    Bygger kontrolteksten fra Blue Prism.

    Teksten bruges til at undgå, at robotten
    behandler tidligere videresendte mails igen.
    """

    return (
        f"Videresend altid til {mailbox} "
        f"<mailto:{mailbox}>  . "
        "Skriv gerne hvem du tror der skulle "
        "have modtaget mailen og hvorfor."
    )


# -------------------------------------------------
# REGLER
# -------------------------------------------------

RULE_STATUS_ACTIVE = "Aktiv"

RULE_STATUS_INACTIVE = "Inaktiv"

RULE_STATUS_TEST = "Test"

RULE_TYPE_MAIN = "Hovedregel"

RULE_TYPE_ADDITIONAL = "Tillægsregel"

RULE_FORWARD_MINIMUM_POINTS = 300


# -------------------------------------------------
# AZURE
# -------------------------------------------------

MAX_AZURE_ATTEMPTS_PER_ITEM = 50

