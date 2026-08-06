"""
CENTRAL PROCESKONFIGURATION

Denne fil styrer:

- testpostkasse
- produktionspostkasser
- valg af regelpostkasse under test
- grænse for antal hentede mails
- robotkategorier
- manuel genbehandling
- regelkonstanter
- Azure-grænse


==================================================
PRODUKTION
==================================================

Brug disse værdier i .env:

TEST_MAILBOX_OVERRIDE=
RULE_MAILBOX_OVERRIDE=

Det betyder:

- Mails hentes fra alle aktive postkasser
  i PRODUCTION_MAILBOXES.

- Hver mail vurderes med reglerne for den
  postkasse, som mailen blev hentet fra.


==================================================
TEST MED ÉN TESTPOSTKASSE
==================================================

Eksempel i .env:

TEST_MAILBOX_OVERRIDE=robot-data@haderslev.dk
RULE_MAILBOX_OVERRIDE=jobcenter@haderslev.dk

Det betyder:

- Mails hentes kun fra:
  robot-data@haderslev.dk

- Mails vurderes med reglerne for:
  jobcenter@haderslev.dk


==================================================
VIGTIGT
==================================================

Hvis TEST_MAILBOX_OVERRIDE er udfyldt:

- Processen kører i testtilstand.
- Produktionspostkasserne bruges ikke.
- Der hentes kun fra testpostkassen.

Hvis TEST_MAILBOX_OVERRIDE er tom:

- Processen kører i produktionstilstand.
- Alle aktive produktionspostkasser bruges.

RULE_MAILBOX_OVERRIDE bør normalt kun være
udfyldt under test.

Hvis RULE_MAILBOX_OVERRIDE er udfyldt i
produktion, vil alle mails blive vurderet med
reglerne for den samme postkasse.
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
# TESTTILSTAND
# -------------------------------------------------

# Testtilstand er aktiv, når der er angivet
# en testpostkasse i TEST_MAILBOX_OVERRIDE.

TEST_MODE = bool(
    TEST_MAILBOX_OVERRIDE
)


# -------------------------------------------------
# POSTKASSEMODEL
# -------------------------------------------------

@dataclass(frozen=True)
class MailboxConfig:
    """
    Konfiguration for én Outlook-postkasse.

    address:
        Postkassens mailadresse.

    folder:
        Outlook-mappen der hentes mails fra.

    enabled:
        Bestemmer om postkassen er aktiv.
    """

    address: str

    folder: str = "inbox"

    enabled: bool = True


# -------------------------------------------------
# PRODUKTIONSPOSTKASSER
# -------------------------------------------------

# Tilføj alle produktionspostkasser her.
#
# enabled=True:
#     Postkassen bruges i produktion.
#
# enabled=False:
#     Postkassen springes over.

PRODUCTION_MAILBOXES = [
    MailboxConfig(
        address="jobcenter@haderslev.dk",
        folder="inbox",
        enabled=True,
    ),

    MailboxConfig(
        address="post@haderslev.dk",
        folder="inbox",
        enabled=True,
    ),
]


# -------------------------------------------------
# VÆLG AKTIVE POSTKASSER
# -------------------------------------------------

if TEST_MODE:
    # TEST:
    #
    # Når TEST_MAILBOX_OVERRIDE er udfyldt,
    # bruges kun den angivne testpostkasse.
    #
    # Produktionspostkasserne bruges ikke.

    MAILBOXES = [
        MailboxConfig(
            address=TEST_MAILBOX_OVERRIDE,
            folder="inbox",
            enabled=True,
        ),
    ]

else:
    # PRODUKTION:
    #
    # Når TEST_MAILBOX_OVERRIDE er tom,
    # bruges alle aktive produktionspostkasser.

    MAILBOXES = [
        mailbox
        for mailbox in PRODUCTION_MAILBOXES
        if mailbox.enabled
    ]


# -------------------------------------------------
# KONTROLLÉR KONFIGURATION
# -------------------------------------------------

if TEST_MODE and not RULE_MAILBOX_OVERRIDE:
    raise ValueError(
        "Processen kører i testtilstand, men "
        "RULE_MAILBOX_OVERRIDE er tom. "
        "Angiv hvilken produktionspostkasses "
        "regler testmailen skal vurderes med."
    )

if not TEST_MODE and RULE_MAILBOX_OVERRIDE:
    raise ValueError(
        "Processen kører i produktionstilstand, "
        "men RULE_MAILBOX_OVERRIDE er udfyldt. "
        "Tøm RULE_MAILBOX_OVERRIDE i .env, så "
        "hver postkasse bruger sine egne regler."
    )

if not MAILBOXES:
    raise ValueError(
        "Der er ingen aktive postkasser."
    )


# -------------------------------------------------
# MAKSIMALT ANTAL MAILS
# -------------------------------------------------

# Værdien gælder pr. postkasse.
#
# Eksempel:
#
# MAIL_LIMIT_PER_MAILBOX = 5
#
# betyder:
#
# - højst 5 mails fra jobcenter
# - højst 5 mails fra post
#
# None betyder, at alle tilgængelige mails hentes.

if TEST_MODE:
    # Brug et lavt antal under test.

    MAIL_LIMIT_PER_MAILBOX = 20

else:
    # Produktionsgrænse pr. postkasse.

    MAIL_LIMIT_PER_MAILBOX = 200


# -------------------------------------------------
# MANUEL GENBEHANDLING
# -------------------------------------------------

# Hvis brugeren sætter denne Outlook-kategori,
# må mailen behandles igen, selv om:
#
# - mailen har robotkategorier
# - mailen indeholder tidligere robottekst
# - mailen tidligere findes som Completed
#
# Kategorien fjernes ikke ved queue-oprettelse.
# Ved vellykket behandling overskrives alle
# kategorier med det nye resultat.

ROBOT_REPROCESS_CATEGORY = (
    "Robot genbehandel"
)


# -------------------------------------------------
# ROBOTKATEGORIER
# -------------------------------------------------

# Mails med disse kategorier betragtes som
# tidligere behandlet.
#
# Kategorien Robot genbehandel kontrolleres
# altid før disse markører.

ROBOT_CATEGORY_MARKERS = (
    "Point: ",
    "Test",
    "Nr:",
)


# -------------------------------------------------
# TIDLIGERE ROBOTVIDERESENDELSE
# -------------------------------------------------

ROBOT_RULE_NUMBER_MARKER = (
    "Regelnr:"
)


def build_previous_forward_text(
    mailbox,
):
    """
    Bygger den gamle kontroltekst fra Blue Prism.

    Funktionen bruges til at genkende mails,
    som tidligere er videresendt af robotten.
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

RULE_STATUS_ACTIVE = (
    "Aktiv"
)

RULE_STATUS_INACTIVE = (
    "Inaktiv"
)

RULE_STATUS_TEST = (
    "Test"
)

RULE_TYPE_MAIN = (
    "Hovedregel"
)

RULE_TYPE_ADDITIONAL = (
    "Tillægsregel"
)

RULE_FORWARD_MINIMUM_POINTS = 300


# -------------------------------------------------
# AZURE
# -------------------------------------------------

# Azure bruges kun til understøttede billedfiler.
#
# PDF-filer læses kun med pypdf og sendes
# aldrig til Azure Vision.

MAX_AZURE_ATTEMPTS_PER_ITEM = 50


# -------------------------------------------------
# VIS AKTIV KONFIGURATION
# -------------------------------------------------

def print_active_configuration():
    """
    Udskriver den aktive proceskonfiguration.

    Funktionen kan kaldes fra main.py, hvis den
    aktive konfiguration skal ses i terminalen.
    """

    print()
    print("===================================")
    print("AKTIV PROCESKONFIGURATION")
    print("===================================")

    if TEST_MODE:
        print("Tilstand: TEST")
        print(
            "Mails hentes fra:",
            TEST_MAILBOX_OVERRIDE,
        )
        print(
            "Regler bruges for:",
            RULE_MAILBOX_OVERRIDE,
        )

    else:
        print("Tilstand: PRODUKTION")
        print(
            "Regeloverride: Ingen"
        )
        print(
            "Aktive postkasser:"
        )

        for mailbox in MAILBOXES:
            print(
                "-",
                mailbox.address,
            )

    print(
        "Mailgrænse pr. postkasse:",
        MAIL_LIMIT_PER_MAILBOX,
    )

    print("===================================")
    print()