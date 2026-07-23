"""
AZURE VISION ADAPTER

Wrapper omkring AzureVisionClient.

Bruges af laes_filindhold.py.

Credential håndteres automatisk via:

API_AZURE_VISION

i Automation Server.
"""
from q_azure_vision_api.api_client import (
AzureVisionClient,
)


# -------------------------------------------------
# ER AZURE KLAR?
# -------------------------------------------------

def azure_vision_is_configured():
    """
    Returnerer True.

    Azure er aktiv i denne version.
    """

    return True


# -------------------------------------------------
# LÆS FIL MED AZURE
# -------------------------------------------------

def read_file_with_azure_vision(
    filename,
    file_bytes,
):
    """
    Læser fil med Azure Vision.

    Returnerer ren tekst.
    """

    client = AzureVisionClient()

    result = client.read_image(
        file_bytes=file_bytes,
        language="da",
    )

    text = (
        AzureVisionClient.extract_text(
            result
        )
    )

    return text