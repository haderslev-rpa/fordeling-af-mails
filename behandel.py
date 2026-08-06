"""
BEHANDLING AF ÉT QUEUE-ITEM

Denne fil følger Haderslevs standard for behandel.py.

Procesrækkefølge:

1.0 - Mail hentet
2.0 - Vedhæftninger hentet
3.0 - Filindhold læst
3.2 - Maksimalt antal Azure-forsøg nået
4.0 - Regler vurderet
5.0 - Mail behandlet

VIGTIGT:

- Mailens body gemmes ikke i item.data.
- Vedhæftningernes bytes gemmes ikke i item.data.
- Udlæst dokumenttekst gemmes ikke i item.data.
- CPR-nummeret gemmes aldrig.
- Kun den valgte regels resultat gemmes i box.
- PDF-filer læses kun med pypdf.
- Billedfiler kan læses med Azure Vision.
- Regler under 300 point overskriver mailens kategorier.
- Testregler overskriver mailens kategorier.
- Aktive regler med mindst 300 point kan videresende
  eller flytte mailen.
- Outlook-handlingen udføres i MAIL_BEHANDLET.
- behandel.py sætter ikke status.
- behandel.py sætter ikke status_code.
- main.py sætter status og status_code til sidst.
"""


async def behandel_page(
    item,
    all_rules,
    debug=False,
):
    """
    Behandler ét queue-item.

    item:
        Det aktuelle Automation Server-item.

    all_rules:
        Alle regler fra alle Excel-regelark.
        Reglerne er hentet én gang i main.py.

    debug:
        Styrer ekstra logning.
    """

    from q_haderslev_vbo.automation_server.ats_update_item_data import (
        update_item_data,
    )

    from q_haderslev_vbo.automation_server.ats_find_state import (
        find_state,
    )

    from q_outlook_api.functionality.mail_api import (
        MailNotFoundError,
        get_attachments,
        get_mails,
    )

    from hent_mailregler_fra_sharepoint import (
        get_rules_for_mailbox,
    )

    from laes_filindhold import (
        read_attachment_contents,
    )

    from udfoer_mailhandling import (
        build_comparison_message,
        execute_mail_action,
    )

    from vurder_mail_mod_regler import (
        select_winning_rule,
    )

    import logging

    logger = logging.getLogger(__name__)

    data = item.data


    # ==========================================================
    # 🧠 STATES
    # ==========================================================

    class States:
        MAIL_HENTET = (
            "1.0 - Mail hentet"
        )

        VEDHAEFTNINGER_HENTET = (
            "2.0 - Vedhæftninger hentet"
        )

        FILINDHOLD_LAEST = (
            "3.0 - Filindhold læst"
        )

        AZURE_FORSOEG_NAAET = (
            "3.2 - Maksimalt antal Azure-forsøg nået"
        )

        REGLER_VURDERET = (
            "4.0 - Regler vurderet"
        )

        MAIL_BEHANDLET = (
            "5.0 - Mail behandlet"
        )


    # ==========================================================
    # 🔁 HELPERS
    # ==========================================================

    def mangler_state(
        state,
        step,
    ):
        """
        Kontrollerer om en state mangler.

        Hvis state allerede findes, bliver
        procestrinnet logget som sprunget over.
        """

        states = data.get(
            "state",
            [],
        )

        match = next(
            (
                existing_state
                for existing_state in states
                if state in existing_state
            ),
            None,
        )

        if match:
            log_step(
                step,
                f'Skip "{match}"',
            )

            return False

        return True


    def set_state(state):
        """
        Tilføjer state gennem standardfunktionen.
        """

        update_item_data(
            data,
            item=item,
            state=state,
        )


    def log_step(
        step,
        text,
    ):
        """
        Skriver en ensartet logbesked.
        """

        logger.info(
            f"[{step}] {text}"
        )


    # ==========================================================
    # INPUT FRA ITEM.DATA
    # ==========================================================

    mailbox = data["box"]["mail"][
        "mailbox"
    ]

    message_id = data["box"]["mail"][
        "message_id"
    ]


    # ==========================================================
    # KONTROLLÉR OM MAILEN ALLEREDE ER BEHANDLET
    # ==========================================================

    if find_state(
        data,
        States.MAIL_BEHANDLET,
    ):
        log_step(
            "MAIL_BEHANDLET",
            f'Skip "{States.MAIL_BEHANDLET}"',
        )

        comparison_message = (
            data["box"].get(
                "comparison_message"
            )
            or "Mailen var allerede behandlet"
        )

        return {
            "already_processed": True,
            "mail_found": None,
            "rules_checked": 0,
            "attachments_read": 0,
            "file_results_count": 0,
            "rule_result": None,
            "action_result": {
                "action": data[
                    "box"
                ].get("action"),
                "destination": data[
                    "box"
                ].get("destination"),
                "status_message": (
                    "Mailen var allerede behandlet"
                ),
            },
            "comparison_message": (
                comparison_message
            ),
            "status_message": (
                "Mailen var allerede behandlet"
            ),
        }


    # ==========================================================
    # FILTRÉR REGLER TIL POSTKASSEN
    # ==========================================================

    rules = get_rules_for_mailbox(
        all_rules=all_rules,
        item_mailbox=mailbox,
    )

    logger.info(
        "[REGLER] %s regler valgt for %s",
        len(rules),
        mailbox,
    )


    # ==========================================================
    # MELLEMRESULTATER I HUKOMMELSEN
    # ==========================================================

    mail = None

    attachments = []

    file_results = []

    rule_result = None

    action_result = None

    comparison_message = "Completed"


    # ==========================================================
    step = "MAIL_HENTET"
    # ==========================================================
    state = getattr(
        States,
        step,
    )

    if mangler_state(
        state,
        step,
    ):

        log_step(
            step,
            "Start",
        )

        try:
            mails = get_mails(
                user_mail=mailbox,
                message_id=message_id,
                include_attachments=True,
                get_inline=False,
                prefer_plain_text=True,
            )

        except MailNotFoundError:
            mails = []

        if not mails:
            data["box"]["action"] = (
                "mail_not_found"
            )

            data["box"]["destination"] = None

            comparison_message = (
                "Mailen findes ikke længere "
                "i Outlook"
            )

            data["box"][
                "comparison_message"
            ] = comparison_message

            log_step(
                step,
                comparison_message,
            )

            update_item_data(
                data,
                item=item,
            )

            set_state(
                States.MAIL_BEHANDLET
            )

            return {
                "already_processed": False,
                "mail_found": False,
                "rules_checked": 0,
                "attachments_read": 0,
                "file_results_count": 0,
                "rule_result": None,
                "action_result": {
                    "action": (
                        "mail_not_found"
                    ),
                    "destination": None,
                    "status_message": (
                        comparison_message
                    ),
                },
                "comparison_message": (
                    comparison_message
                ),
                "status_message": (
                    comparison_message
                ),
            }

        mail = mails[0]

        log_step(
            step,
            (
                "Mail hentet: "
                f'{mail.get("subject") or ""}'
            ),
        )

        set_state(state)

    else:
        # Mailens body gemmes ikke i item.data.
        #
        # Mailen skal derfor hentes igen til
        # hukommelsen ved en genkørsel.

        try:
            mails = get_mails(
                user_mail=mailbox,
                message_id=message_id,
                include_attachments=True,
                get_inline=False,
                prefer_plain_text=True,
            )

        except MailNotFoundError:
            mails = []

        if not mails:
            data["box"]["action"] = (
                "mail_not_found"
            )

            data["box"]["destination"] = None

            comparison_message = (
                "Mailen findes ikke længere "
                "i Outlook"
            )

            data["box"][
                "comparison_message"
            ] = comparison_message

            log_step(
                step,
                comparison_message,
            )

            update_item_data(
                data,
                item=item,
            )

            set_state(
                States.MAIL_BEHANDLET
            )

            return {
                "already_processed": False,
                "mail_found": False,
                "rules_checked": 0,
                "attachments_read": 0,
                "file_results_count": 0,
                "rule_result": None,
                "action_result": {
                    "action": (
                        "mail_not_found"
                    ),
                    "destination": None,
                    "status_message": (
                        comparison_message
                    ),
                },
                "comparison_message": (
                    comparison_message
                ),
                "status_message": (
                    comparison_message
                ),
            }

        mail = mails[0]


    # ==========================================================
    step = "VEDHAEFTNINGER_HENTET"
    # ==========================================================
    state = getattr(
        States,
        step,
    )

    if mangler_state(
        state,
        step,
    ):

        log_step(
            step,
            "Start",
        )

        attachments = get_attachments(
            user_mail=mailbox,
            message_id=message_id,
            get_inline=False,
        )

        log_step(
            step,
            (
                f"{len(attachments)} "
                "vedhæftninger hentet"
            ),
        )

        if debug:
            for attachment in attachments:
                log_step(
                    step,
                    (
                        "Fil: "
                        f'{attachment.get("name")}. '
                        "Type: "
                        f'{attachment.get("attachment_type")}. '
                        "Størrelse: "
                        f'{attachment.get("size")}.'
                    ),
                )

        set_state(state)

    else:
        # Attachment-bytes gemmes ikke i item.data.
        #
        # Vedhæftningerne skal derfor hentes igen
        # til hukommelsen ved en genkørsel.

        attachments = get_attachments(
            user_mail=mailbox,
            message_id=message_id,
            get_inline=False,
        )


    # ==========================================================
    step = "FILINDHOLD_LAEST"
    # ==========================================================
    state = getattr(
        States,
        step,
    )

    if mangler_state(
        state,
        step,
    ):

        log_step(
            step,
            "Start",
        )

        file_results = (
            read_attachment_contents(
                attachments=attachments,
                item=item,
            )
        )

        log_step(
            step,
            (
                f"{len(file_results)} "
                "filer gennemgået"
            ),
        )

        if debug:
            for file_result in file_results:
                log_step(
                    step,
                    (
                        "Fil: "
                        f'{file_result.get("name")}. '
                        "Metode: "
                        f'{file_result.get("read_method")}. '
                        "Success: "
                        f'{file_result.get("success")}. '
                        "Azure: "
                        f'{file_result.get("azure_used")}.'
                    ),
                )

                if file_result.get("error"):
                    log_step(
                        step,
                        (
                            "Filfejl: "
                            f'{file_result.get("error")}'
                        ),
                    )

        set_state(state)

    else:
        # pypdf skal læse PDF-filerne igen ved
        # hvert worker-forsøg.
        #
        # Billeder kan læses igen med Azure,
        # hvis Azure-grænsen tillader det.
        #
        # Filtekst gemmes ikke i item.data.

        file_results = (
            read_attachment_contents(
                attachments=attachments,
                item=item,
            )
        )


    # ==========================================================
    # KONTROLLÉR AZURE-GRÆNSE
    # ==========================================================

    azure_limit_reached = any(
        (
            file_result.get(
                "azure_limit_reached",
                False,
            )
            or file_result.get(
                "read_method"
            ) == "azure_limit_reached"
        )
        for file_result in file_results
    )

    if azure_limit_reached:

        # ==========================================================
        step = "AZURE_FORSOEG_NAAET"
        # ==========================================================
        state = getattr(
            States,
            step,
        )

        if mangler_state(
            state,
            step,
        ):

            log_step(
                step,
                (
                    "Maksimalt antal "
                    "Azure-forsøg nået"
                ),
            )

            set_state(state)


    # ==========================================================
    step = "REGLER_VURDERET"
    # ==========================================================
    state = getattr(
        States,
        step,
    )

    if mangler_state(
        state,
        step,
    ):

        log_step(
            step,
            "Start",
        )

        rule_result = select_winning_rule(
            rules=rules,
            mail=mail,
            file_results=file_results,
        )

        data["box"]["subject_points"] = (
            rule_result.get(
                "subject_points",
                0,
            )
        )

        data["box"]["subject_match"] = (
            rule_result.get(
                "subject_match",
                "",
            )
        )

        data["box"]["message_points"] = (
            rule_result.get(
                "message_points",
                0,
            )
        )

        data["box"]["message_match"] = (
            rule_result.get(
                "message_match",
                "",
            )
        )

        data["box"]["sender_points"] = (
            rule_result.get(
                "sender_points",
                0,
            )
        )

        data["box"]["sender_match"] = (
            rule_result.get(
                "sender_match",
                "",
            )
        )

        data["box"]["file_name_points"] = (
            rule_result.get(
                "file_name_points",
                0,
            )
        )

        data["box"]["file_name_match"] = (
            rule_result.get(
                "file_name_match",
                "",
            )
        )

        data["box"][
            "file_content_points"
        ] = rule_result.get(
            "file_content_points",
            0,
        )

        data["box"][
            "file_content_match"
        ] = rule_result.get(
            "file_content_match",
            "",
        )

        data["box"]["cpr_points"] = (
            rule_result.get(
                "cpr_points",
                0,
            )
        )

        data["box"]["cpr_found"] = bool(
            rule_result.get(
                "cpr_found",
                False,
            )
        )

        data["box"]["rule_number"] = (
            rule_result.get(
                "rule_number"
            )
        )

        data["box"]["total_points"] = (
            rule_result.get(
                "total_points",
                0,
            )
        )

        log_step(
            step,
            (
                "Regelnummer: "
                f'{data["box"]["rule_number"]}. '
                "Total point: "
                f'{data["box"]["total_points"]}.'
            ),
        )

        update_item_data(
            data,
            item=item,
        )

        set_state(state)

    else:
        # Regelresultatet beregnes igen til
        # Outlook-handlingen.
        #
        # Kun den valgte regels felter er
        # gemt i box.
        #
        # Første regel med mindst 300 point
        # stopper regelgennemgangen.
        #
        # Hvis ingen regel når 300 point,
        # bruges reglen med flest point.
        #
        # Ved pointlighed bruges den første regel.

        rule_result = select_winning_rule(
            rules=rules,
            mail=mail,
            file_results=file_results,
        )


    # ==========================================================
    step = "MAIL_BEHANDLET"
    # ==========================================================
    state = getattr(
        States,
        step,
    )

    if mangler_state(
        state,
        step,
    ):

        log_step(
            step,
            "Start",
        )

        # Outlook-handlingen udføres her.
        #
        # Under 300 point:
        #     Alle kategorier overskrives.
        #     Mailen videresendes eller flyttes ikke.
        #
        # Testregel:
        #     Alle kategorier overskrives.
        #     Mailen videresendes eller flyttes ikke.
        #
        # Aktiv regel med mindst 300 point:
        #     Mailen videresendes eller flyttes.

        action_result = execute_mail_action(
            mail=data["box"]["mail"],
            rule_result=rule_result,
        )

        comparison_message = (
            build_comparison_message(
                rule_result=rule_result,
                action_result=action_result,
            )
        )

        data["box"][
            "comparison_message"
        ] = comparison_message

        data["box"]["action"] = (
            action_result.get(
                "action"
            )
        )

        data["box"]["destination"] = (
            action_result.get(
                "destination"
            )
        )

        log_step(
            step,
            (
                "Handling: "
                f'{data["box"]["action"]}. '
                "Destination: "
                f'{data["box"]["destination"]}. '
                "Message: "
                f"{comparison_message}."
            ),
        )

        update_item_data(
            data,
            item=item,
        )

        # State sættes først efter:
        #
        # 1. Outlook-handlingen er gennemført.
        # 2. Resultatet er gemt i item.data.
        #
        # Hvis Outlook-handlingen fejler,
        # bliver state derfor ikke sat.

        set_state(state)

    else:
        action_result = {
            "action": data[
                "box"
            ].get("action"),
            "destination": data[
                "box"
            ].get("destination"),
            "status_message": (
                "Mailen var allerede behandlet"
            ),
        }

        comparison_message = (
            data["box"].get(
                "comparison_message"
            )
            or "Mailen var allerede behandlet"
        )


    # ==========================================================
    # RETURNÉR RESULTAT TIL MAIN.PY
    # ==========================================================

    return {
        "already_processed": False,

        "mail_found": True,

        "rules_checked": len(
            rules
        ),

        "attachments_read": len(
            attachments
        ),

        "file_results_count": len(
            file_results
        ),

        "rule_result": (
            rule_result
        ),

        "action_result": (
            action_result
        ),

        "comparison_message": (
            comparison_message
        ),

        # main.py bruger denne værdi
        # til den afsluttende status.
        "status_message": (
            action_result.get(
                "status_message"
            )
            or "Mail behandlet"
        ),
    }