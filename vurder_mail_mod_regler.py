"""
REGELMOTOR TIL FORDELING AF MAILS

Denne fil:

1. Modtager reglerne for den aktuelle postkasse.
2. Søger i mailens emne.
3. Søger i mailens body.
4. Søger i afsenderens mailadresse.
5. Søger i alle filnavne.
6. Søger i alt udlæst filindhold.
7. Kontrollerer om CPR findes.
8. Beregner point for alle regler.
9. Returnerer kun den vindende regel.

VIGTIGT:

## betyder ELLER.

&& betyder OG.

CPR-nummeret returneres aldrig.
Kun cpr_found som True eller False returneres.

Mail-body og filindhold bruges kun i hukommelsen.
"""

import re


# -------------------------------------------------
# KONSTANTER
# -------------------------------------------------

ACTIVE_MINIMUM_POINTS = 300


# -------------------------------------------------
# NORMALISÉR TEKST
# -------------------------------------------------

def normalize_text(value):
    """
    Normaliserer tekst til sammenligning.

    casefold gør sammenligningen uafhængig
    af store og små bogstaver.
    """

    if value is None:
        return ""

    return str(value).casefold()


# -------------------------------------------------
# OPDEL ELLER-MØNSTRE
# -------------------------------------------------

def split_or_patterns(pattern):
    """
    Opdeler et mønster ved ##.

    Eksempel:

    DP 235##DP 211

    bliver til:

    [
        "DP 235",
        "DP 211"
    ]
    """

    return [
        part.strip()
        for part in str(pattern or "").split("##")
        if part.strip()
    ]


# -------------------------------------------------
# OPDEL OG-MØNSTRE
# -------------------------------------------------

def split_and_patterns(pattern):
    """
    Opdeler et mønster ved &&.

    Eksempel:

    Brev&&Aktindsigt

    betyder, at begge tekster skal findes.
    """

    return [
        part.strip()
        for part in str(pattern or "").split("&&")
        if part.strip()
    ]


# -------------------------------------------------
# MATCH MØNSTER
# -------------------------------------------------

def match_pattern(
    text,
    pattern,
):
    """
    Søger efter et Excel-mønster i en tekst.

    ## betyder ELLER.
    && betyder OG.

    Returnerer:

    matched:
        True eller False.

    matched_text:
        Den del af reglen, som gav match.
    """

    pattern_text = str(
        pattern or ""
    ).strip()

    if not pattern_text:
        return False, ""

    searchable_text = normalize_text(
        text
    )

    or_patterns = split_or_patterns(
        pattern_text
    )

    for or_pattern in or_patterns:
        and_patterns = split_and_patterns(
            or_pattern
        )

        if not and_patterns:
            continue

        all_patterns_found = all(
            normalize_text(and_pattern)
            in searchable_text
            for and_pattern in and_patterns
        )

        if all_patterns_found:
            return True, or_pattern

    return False, ""


# -------------------------------------------------
# BEREGN POINT FOR ÉT FELT
# -------------------------------------------------

def calculate_field_points(
    text,
    pattern,
    configured_points,
):
    """
    Beregner point for ét regelområde.

    Hvis mønstret findes:
        Returnér de konfigurerede point.

    Hvis mønstret ikke findes:
        Returnér 0 point.
    """

    matched, matched_text = match_pattern(
        text=text,
        pattern=pattern,
    )

    if not matched:
        return 0, ""

    return (
        int(configured_points or 0),
        matched_text,
    )


# -------------------------------------------------
# CPR-MØNSTER
# -------------------------------------------------

CPR_PATTERN = re.compile(
    r"(?<!\d)"
    r"(?:0[1-9]|[12]\d|3[01])"
    r"(?:0[1-9]|1[0-2])"
    r"\d{2}"
    r"[- ]?"
    r"\d{4}"
    r"(?!\d)"
)


# -------------------------------------------------
# FIND CPR
# -------------------------------------------------

def contains_cpr(*texts):
    """
    Kontrollerer om et CPR-lignende mønster findes.

    Funktionen returnerer kun True eller False.
    Selve CPR-nummeret returneres aldrig.
    """

    for text in texts:
        if CPR_PATTERN.search(
            str(text or "")
        ):
            return True

    return False


# -------------------------------------------------
# BEREGN CPR-POINT
# -------------------------------------------------

def calculate_cpr_points(
    configured_points,
    cpr_found,
):
    """
    Beregner CPR-point.

    Eksempel med positivt point:

    configured_points = 300
    CPR findes
    resultat = 300

    Eksempel med negativt point:

    configured_points = -1000
    CPR findes
    resultat = -1000

    Det gør det muligt at forhindre en regel i
    at vinde, når mailen indeholder CPR.

    Hvis CPR ikke findes, gives 0 point.
    """

    configured_points = int(
        configured_points or 0
    )

    if not cpr_found:
        return 0

    return configured_points


# -------------------------------------------------
# SAMMENFØJ FILNAVNE
# -------------------------------------------------

def combine_file_names(mail):
    """
    Samler alle filnavne til søgning.

    Filnavnene bliver ikke ændret i item.data.
    """

    attachment_names = (
        mail.get("attachment_names")
        or []
    )

    return "\n".join(
        str(filename)
        for filename in attachment_names
    )


# -------------------------------------------------
# SAMMENFØJ FILINDHOLD
# -------------------------------------------------

def combine_file_contents(file_results):
    """
    Samler udlæst filtekst i hukommelsen.

    Teksten gemmes ikke i item.data.
    """

    file_texts = []

    for file_result in file_results or []:
        text = file_result.get(
            "text"
        )

        if text:
            file_texts.append(
                str(text)
            )

    return "\n".join(
        file_texts
    )


# -------------------------------------------------
# HENT MAILENS BODY
# -------------------------------------------------

def get_mail_body(mail):
    """
    Returnerer mailens body.

    q-outlook-api kan returnere body i både
    body og body_text.
    """

    return (
        mail.get("body")
        or mail.get("body_text")
        or ""
    )


# -------------------------------------------------
# HENT MAILENS AFSENDER
# -------------------------------------------------

def get_mail_sender(mail):
    """
    Returnerer afsenderens mailadresse.
    """

    return (
        mail.get("sender_address")
        or mail.get("from_email")
        or ""
    )


# -------------------------------------------------
# FIND KRÆVEDE POINT
# -------------------------------------------------

def get_required_points(rule):
    """
    Finder grænsen for en regel.

    Aktive regler kræver mindst 300 point.

    Testregler bruger Max Point, hvis den er
    større end 0. Det gør det muligt at teste
    regler med eksempelvis 200 eller 290 point.
    """

    rule_status = str(
        rule.status or ""
    ).casefold()

    max_points = int(
        rule.max_points or 0
    )

    if (
        rule_status == "test"
        and max_points > 0
    ):
        return max_points

    return ACTIVE_MINIMUM_POINTS


# -------------------------------------------------
# BEREGN ÉN REGEL
# -------------------------------------------------

def evaluate_rule(
    rule,
    mail,
    file_names_text,
    file_content_text,
    cpr_found,
):
    """
    Beregner alle point for én regel.
    """

    mail_subject = (
        mail.get("subject")
        or ""
    )

    mail_body = get_mail_body(
        mail
    )

    mail_sender = get_mail_sender(
        mail
    )

    (
        subject_points,
        subject_match,
    ) = calculate_field_points(
        text=mail_subject,
        pattern=rule.subject_pattern,
        configured_points=(
            rule.subject_points
        ),
    )

    (
        message_points,
        message_match,
    ) = calculate_field_points(
        text=mail_body,
        pattern=rule.message_pattern,
        configured_points=(
            rule.message_points
        ),
    )

    (
        sender_points,
        sender_match,
    ) = calculate_field_points(
        text=mail_sender,
        pattern=rule.sender_pattern,
        configured_points=(
            rule.sender_points
        ),
    )

    (
        file_name_points,
        file_name_match,
    ) = calculate_field_points(
        text=file_names_text,
        pattern=rule.file_name_pattern,
        configured_points=(
            rule.file_name_points
        ),
    )

    (
        file_content_points,
        file_content_match,
    ) = calculate_field_points(
        text=file_content_text,
        pattern=rule.file_content_pattern,
        configured_points=(
            rule.file_content_points
        ),
    )

    cpr_points = calculate_cpr_points(
        configured_points=(
            rule.cpr_points
        ),
        cpr_found=cpr_found,
    )

    total_points = sum(
        (
            subject_points,
            message_points,
            sender_points,
            file_name_points,
            file_content_points,
            cpr_points,
        )
    )

    required_points = get_required_points(
        rule
    )

    qualifies = (
        total_points >= required_points
    )

    return {
        "excel_row_number": (
            rule.excel_row_number
        ),

        "subject_points": subject_points,
        "subject_match": subject_match,

        "message_points": message_points,
        "message_match": message_match,

        "sender_points": sender_points,
        "sender_match": sender_match,

        "file_name_points": (
            file_name_points
        ),
        "file_name_match": (
            file_name_match
        ),

        "file_content_points": (
            file_content_points
        ),
        "file_content_match": (
            file_content_match
        ),

        "cpr_points": cpr_points,
        "cpr_found": cpr_found,

        "rule_number": (
            rule.rule_number
        ),
        "total_points": (
            total_points
        ),

        "rule_status": (
            rule.status
        ),
        "rule_type": (
            rule.rule_type
        ),
        "destination": (
            rule.destination
        ),

        "required_points": (
            required_points
        ),

        "qualifies": qualifies,
    }


# -------------------------------------------------
# BYG TOMT RESULTAT
# -------------------------------------------------

def build_empty_result(cpr_found):
    """
    Bygger resultatet, når ingen regel matcher.
    """

    return {
        "excel_row_number": None,

        "subject_points": 0,
        "subject_match": "",

        "message_points": 0,
        "message_match": "",

        "sender_points": 0,
        "sender_match": "",

        "file_name_points": 0,
        "file_name_match": "",

        "file_content_points": 0,
        "file_content_match": "",

        "cpr_points": 0,
        "cpr_found": cpr_found,

        "rule_number": None,
        "total_points": 0,

        "rule_status": None,
        "rule_type": None,
        "destination": None,

        "required_points": None,

        "qualifies": False,
    }


# -------------------------------------------------
# VÆLG VINDENDE REGEL
# -------------------------------------------------

def select_winning_rule(
    rules,
    mail,
    file_results,
):
    """
    Beregner alle regler for den aktuelle postkasse.

    Kun regler, som når pointgrænsen, kan vinde.

    Vinderen vælges således:

    1. Højeste total point.
    2. Ved samme point vinder første Excel-række.

    Kun den vindende regel returneres.
    """

    file_names_text = combine_file_names(
        mail
    )

    file_content_text = (
        combine_file_contents(
            file_results
        )
    )

    mail_body = get_mail_body(
        mail
    )

    cpr_found = contains_cpr(
        mail.get("subject"),
        mail_body,
        file_content_text,
    )

    qualifying_results = []

    for rule in rules or []:
        result = evaluate_rule(
            rule=rule,
            mail=mail,
            file_names_text=(
                file_names_text
            ),
            file_content_text=(
                file_content_text
            ),
            cpr_found=cpr_found,
        )

        if result["qualifies"]:
            qualifying_results.append(
                result
            )

    if not qualifying_results:
        return build_empty_result(
            cpr_found=cpr_found
        )

    qualifying_results.sort(
        key=lambda result: (
            -result["total_points"],
            result["excel_row_number"],
        )
    )

    return qualifying_results[0]