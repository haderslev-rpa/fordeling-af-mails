"""
AZURE VISION ADAPTER

Denne fil er et tilkoblingspunkt til jeres eksisterende
Azure Vision-kode.

Azure Vision er deaktiveret i denne første worker-version.
PDF-filer med tekstlag læses stadig med pypdf.
"""

# -------------------------------------------------
# ER AZURE KLAR?
# -------------------------------------------------

def azure_vision_is_configured():
    """
    Returnerer False indtil Azure-kaldet er koblet på.
    """

    return False


# -------------------------------------------------
# LÆS FIL MED AZURE
# -------------------------------------------------

def read_file_with_azure_vision(
    filename,
    file_bytes,
):
    """
    Skal senere kalde jeres eksisterende
    Azure Vision-funktion.

    Azure-objectet skal selv håndtere
    konti, kvoter og genforsøg.
    """

    raise NotImplementedError(
        "Azure Vision-adapteren er ikke koblet på."
    )