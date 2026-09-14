"""SQL queries for the Casewise TAT report.

QUERY_BASE_DATA is a template (not a ready-to-run string) - it must be built
via build_base_data_query(), which fills in the lookback window and an
optional client scope. This is necessary for performance: "all clients,
last N months" is genuinely millions of rows on this database, so every
caller should scope to a specific client (or a small set) wherever possible.
QUERY_CLIENT_LIST is a fast, separate lookup used to power the client
picker without touching the large case tables at all.
"""
from typing import Iterable, Optional

QUERY_CLIENT_LIST = """
SELECT emc.company_id, emc.company_name, ec.client_external_id
FROM ec_client ec
JOIN ec_master_company emc ON emc.company_id = ec.client_id
ORDER BY emc.company_name
"""

_QUERY_BASE_DATA_TEMPLATE = """

WITH base AS (
    SELECT
        TRIM(
            CONCAT(
                CASE
                    WHEN ecc.first_name IS NULL OR ecc.first_name = '' THEN ''
                    ELSE ecc.first_name
                END,
                CASE
                    WHEN ecc.middle_name IS NULL OR ecc.middle_name = '' THEN ''
                    ELSE CONCAT(' ', ecc.middle_name)
                END,
                CASE
                    WHEN ecc.last_name IS NULL OR ecc.last_name = '' THEN ''
                    ELSE CONCAT(' ', ecc.last_name)
                END
            )
        ) AS candidate_name,
        emc.company_name,
        ec.client_external_id,
        CONCAT(eud1.user_first_name, ' ', eud1.user_last_name) AS `CAT`,
        CONCAT(eud2.user_first_name, ' ', eud2.user_last_name) AS `CAT_TL`,
        CONCAT(eud3.user_first_name, ' ', eud3.user_last_name) AS `Account_Manager`,
        ecm.case_ars_no,
        emcl.office_name AS `Location`,
        ecp.process_name,
        ecm.received_date,
        DATE_FORMAT(ecm.received_date, '%b''%y') AS received_month,
        EXTRACT(YEAR FROM ecm.received_date) AS year,

    (
    SELECT MAX(ch.action_taken_on)
    FROM ec_case_history ch
    LEFT JOIN ec_case_reports ecr
         ON ch.case_id = ecr.case_id
         AND ecr.report_type = 1            -- keep this in JOIN to avoid filtering out rows
    JOIN ec_case_checks ecc
         ON ch.check_id = ecc.case_check_id
    WHERE ch.case_id = ecm.case_id
         AND ch.action_taken IN (

' case Status changed to : Case Insufficient',
' case Status changed to : InSufficient',
'Insuff Raised',
'Marked Insufficient',
'Marked Insufficient (Parallel Research)',
'New Status - Case Insufficient',
'New Status - InSufficient',
'Vendor request closed & Insuff raised',
'Case Insufficient',
'Check Created | Marked Insufficient',
'Check Updated | Marked Insufficient',
'New Status - Case Insuff Updated',
'New Status - New Case | Case Insufficient',
'Check Insuff raised',
'Case level Insuff raised',
'New Status - Case Insufficiency raised',
'Insufficient - Rework on Report',
'Insuff raised Comments',
'Insuff accepted',
'Insuff raised by highlighter'

      )
      AND ecc.check_status <> 9
      AND (
            -- apply only when report exists
            (ecr.report_sent_on IS NOT NULL AND ch.action_taken_on <= ecr.report_sent_on)
            OR
            -- if no report or report_sent_on is NULL → include all rows
            (ecr.report_sent_on IS NULL)
          )
) AS Latest_check_insuff_raised_date,

(
    SELECT MIN(ch.action_taken_on)
    FROM ec_case_history ch
    LEFT JOIN ec_case_reports ecr
         ON ch.case_id = ecr.case_id
         AND ecr.report_type = 1            -- keep this in JOIN to avoid filtering out rows
    JOIN ec_case_checks ecc
         ON ch.check_id = ecc.case_check_id
    WHERE ch.case_id = ecm.case_id
         AND ch.action_taken IN (

' case Status changed to : Case Insufficient',
' case Status changed to : InSufficient',
'Insuff Raised',
'Marked Insufficient',
'Marked Insufficient (Parallel Research)',
'New Status - Case Insufficient',
'New Status - InSufficient',
'Vendor request closed & Insuff raised',
'Case Insufficient',
'Check Created | Marked Insufficient',
'Check Updated | Marked Insufficient',
'New Status - Case Insuff Updated',
'New Status - New Case | Case Insufficient',
'Check Insuff raised',
'Case level Insuff raised',
'New Status - Case Insufficiency raised',
'Insufficient - Rework on Report',
'Insuff raised Comments',
'Insuff accepted',
'Insuff raised by highlighter'

      )
      AND ecc.check_status <> 9
      AND (
            -- apply only when report exists
            (ecr.report_sent_on IS NOT NULL AND ch.action_taken_on <= ecr.report_sent_on)
            OR
            -- if no report or report_sent_on is NULL → include all rows
            (ecr.report_sent_on IS NULL)
          )
) AS first_check_insuff_raised_date,

(
    SELECT MAX(ch.action_taken_on)
    FROM ec_case_history ch
    LEFT JOIN ec_case_reports ecr
         ON ch.case_id = ecr.case_id
         AND ecr.report_type = 1            -- keep this in JOIN to avoid filtering out rows
    JOIN ec_case_checks ecc
         ON ch.check_id = ecc.case_check_id
    WHERE ch.case_id = ecm.case_id
         AND ch.action_taken IN (
            'case Status changed to : Insufficiency Fulfilled',
            'Insufficiency Fulfilled',
            'New Status - Insufficiency Fulfilled',
            'Check Updated during Insuff fulfill',
            'Check Insufficiency Fulfilled',
            'New Status - Case Insufficiency Fulfilled',
            'Partial Insufficiency fulfilled',
            'Insuff Fulfilled'
      )
      AND ecc.check_status <> 9
      AND (
            -- apply only when report exists
            (ecr.report_sent_on IS NOT NULL AND ch.action_taken_on <= ecr.report_sent_on)
            OR
            -- if no report or report_sent_on is NULL → include all rows
            (ecr.report_sent_on IS NULL)
          )
) AS Latest_check_insuff_Fulfilled_date,


(
    SELECT MAX(ch.action_taken_on)
    FROM ec_case_history ch
    LEFT JOIN ec_case_reports ecr
        ON ch.case_id = ecr.case_id
        AND ecr.report_type = 1             -- keep report_type filter in JOIN
    JOIN ec_case_checks ecc
        ON ch.check_id = ecc.case_check_id
    WHERE ch.case_id = ecm.case_id
      AND ch.action_taken LIKE '%Relived%'
      AND ecc.check_status <> 9
      AND (
            -- Apply filter only when report_sent_on is present
            (ecr.report_sent_on IS NOT NULL AND ch.action_taken_on <= ecr.report_sent_on)
            OR
            -- If no report or report_sent_on is NULL → include all rows
            (ecr.report_sent_on IS NULL)
          )
) AS `Go_ahead_Date`,

(
    SELECT MAX(ch.action_taken_on)
    FROM ec_case_history ch
    LEFT JOIN ec_case_reports ecr
        ON ch.case_id = ecr.case_id
        AND ecr.report_type = 1               -- important: keep this in JOIN
    JOIN ec_case_checks ecc
        ON ch.check_id = ecc.case_check_id
    WHERE ch.case_id = ecm.case_id
      AND ch.action_taken LIKE '%Reopen%'
      AND ecc.check_status <> 9
      AND (
            -- apply only when report_sent_on exists
            (ecr.report_sent_on IS NOT NULL AND ch.action_taken_on <= ecr.report_sent_on)
            OR
            -- include all rows if report_sent_on is NULL
            (ecr.report_sent_on IS NULL)
          )
) AS `Reopen_dates`,


        (SELECT report_sent_on
         FROM ec_case_reports ecr
         WHERE ecm.case_id = ecr.case_id
           AND ecr.report_type = 1
         ORDER BY ecr.report_sent_on DESC
         LIMIT 1) AS `Final_Report_Sent`,

        (CASE ecm.case_status
            WHEN '1' THEN 'New (Incomplete)'
            WHEN '2' THEN 'On Hold'
            WHEN '3' THEN 'Insufficient'
            WHEN '4' THEN 'Work in Progress'
            WHEN '5' THEN 'Work in Progress'
            WHEN '6' THEN 'Closed by Client'
            WHEN '7' THEN 'Completed'
            WHEN '8' THEN 'Closed by Authbridge'
            WHEN '9' THEN 'Closed-Case Insufficient'
            WHEN '10' THEN 'HighlighterCase'
            WHEN '11' THEN 'SignOff Pending'
            WHEN '12' THEN 'Pending For Duplicity'
            WHEN '13' THEN 'Duplicity'
            WHEN '14' THEN 'Scrap'
            WHEN '15' THEN 'Excel Pending For Duplicity'
            WHEN '16' THEN 'Escalation Raised'
            WHEN '17' THEN 'Escalation Received'
            WHEN '18' THEN 'Excel Duplicity'
        END) AS `Case_Status`,

        (CASE ecm.case_status
            WHEN 1 THEN 'New (Incomplete)'
            WHEN 2 THEN 'On Hold'
            WHEN 3 THEN 'Insufficient'
            WHEN 4 THEN 'Work in Progress'
            WHEN 5 THEN 'Work in Progress'
            WHEN 6 THEN 'Sent'
            WHEN 7 THEN 'Sent'
        END) AS `Case Type`,
        ROW_NUMBER() OVER(PARTITION BY case_id ORDER BY received_date DESC) AS rn,

        (SELECT report_severity
         FROM ec_case_reports ecr
         WHERE ecm.case_id = ecr.case_id
           AND ecr.report_type = 1
         ORDER BY ecr.report_sent_on DESC
         LIMIT 1) AS `Report Severity`,

        (SELECT report_sent_on
         FROM ec_case_reports ecr
         WHERE ecm.case_id = ecr.case_id
           AND ecr.report_type = 2
         ORDER BY ecr.report_sent_on DESC
         LIMIT 1) AS `Last Additional Report sent date`,

        (SELECT report_severity
         FROM ec_case_reports ecr
         WHERE ecm.case_id = ecr.case_id
           AND ecr.report_type = 2
         ORDER BY ecr.report_sent_on DESC
         LIMIT 1) AS `Last Additional Report severity`,

        (SELECT report_sent_on
         FROM ec_case_reports ecr
         WHERE ecm.case_id = ecr.case_id
           AND ecr.report_type = 0
         ORDER BY ecr.report_sent_on
         LIMIT 1) AS `First Interim Report sent date`,

        (SELECT report_severity
         FROM ec_case_reports ecr
         WHERE ecm.case_id = ecr.case_id
           AND ecr.report_type = 0
         ORDER BY ecr.report_sent_on
         LIMIT 1) AS `First Interim Report severity`,

        (SELECT ecr.report_severity
         FROM ec_case_reports ecr
         WHERE ecm.case_id = ecr.case_id
           AND (ecr.report_type = 2 OR ecr.report_type = 1)
         ORDER BY ecr.report_sent_on DESC
         LIMIT 1) AS `Latest Report Severity`,

        (CASE
            WHEN ectc.TAT_catg IS NULL THEN CONVERT(ecp.TAT_DAYS_TYPE USING utf8mb4)
            ELSE CONVERT(ectc.TAT_catg USING utf8mb4)
        END COLLATE utf8mb4_0900_ai_ci) AS client_tat_type,

        (CASE
            WHEN ectc.actual_tat IS NULL THEN ecp.tat
            ELSE ectc.actual_tat
        END) AS Process_TAT,

        (CASE client_category
	    WHEN 1 THEN 'RED'
	    WHEN 2 THEN 'GREY'
	    WHEN 3 THEN 'GREEN'
	    WHEN 4 THEN 'Amber'
	    WHEN 5 THEN 'Blue'
       END) `Client_Category`
     /*
    CASE
    WHEN ectc.tat_tp = 'Case' THEN ectc.actual_tat
    ELSE MAX(ectc1.actual_tat) OVER (PARTITION BY ecp.process_name)
    END AS Process_TAT

    */

    FROM ec_case_master ecm
    LEFT JOIN ec_case_candidates ecc
    ON ecm.candidate_id = ecc.candidate_id
    LEFT JOIN ec_master_company emc
    ON ecm.client_id = emc.company_id
    LEFT JOIN ec_master_company_locations emcl
    ON ecm.client_office_id = emcl.office_id
    LEFT JOIN ec_client_process ecp
    ON ecm.process_id = ecp.process_id
    LEFT JOIN ec_case_tat_config ectc
    ON (ecp.client_id = ectc.client_id AND ecp.process_id = ectc.process_id)
    LEFT JOIN ec_check_tat_config ectc1 ON ectc.id=ectc1.case_tat_id
    LEFT JOIN ec_client ec ON ecm.client_id = ec.client_id
    LEFT JOIN ec_user_details eud1 ON ec.cat_id = eud1.user_id
    LEFT JOIN ec_user_details eud2 ON ec.cat_tl = eud2.user_id
    LEFT JOIN ec_user_details eud3 ON ec.cat_account_manager = eud3.user_id
    WHERE
    ecm.received_date >= DATE_SUB(CAST(DATE_FORMAT(CURDATE(), '%Y-%m-01') AS DATE), INTERVAL {months_back} MONTH)
    AND ecm.received_date <= CURRENT_DATE
    AND ecc.first_name IS NOT NULL
    AND ecc.first_name NOT LIKE '%Dummy%'
    AND ecc.first_name NOT LIKE '%test%'
    AND ecm.case_status IN (2,3,4,5,6,7)
      -- AND emc.company_id  IN (78698)
    AND emc.company_id  NOT IN (157843,325739,181532,139215,211271,198170,122166,211974,98281,170092)
      -- AND emc.company_id  IN (158686)
      -- AND ecm.case_ars_no='1238-239086'
    {company_filter}
),

T1 AS (
SELECT *,

CASE
    WHEN `first_check_insuff_raised_date` IS NULL
    AND `Latest_check_insuff_raised_date` IS NULL
    AND `Latest_check_insuff_Fulfilled_date` IS NULL
    THEN 'No Insuff'

    WHEN client_category = 1
    AND `Latest_check_insuff_Fulfilled_date` < `Latest_check_insuff_raised_date`
    THEN 'Insuff'

    WHEN `Latest_check_insuff_Fulfilled_date` IS NOT NULL
    THEN 'Fulfill'

    ELSE 'Insuff'
    END AS `Insuff status`


FROM base
WHERE rn = 1),

/* Layer 2: add start_date (so we can reuse it safely later) */
T2 AS (
    SELECT
        T1.*,
        -- MySQL's GREATEST() returns NULL if ANY argument is NULL (unlike
        -- Postgres/Redshift, which ignore NULLs and return the max of the
        -- non-null values) - most cases have no insufficiency history, so
        -- Latest_check_insuff_Fulfilled_date/Go_ahead_Date/Reopen_dates are
        -- NULL and GREATEST() would return NULL for nearly every row
        -- without this COALESCE fallback to received_date (always non-null).
        GREATEST(
            COALESCE(`Latest_check_insuff_Fulfilled_date`, received_date),
            received_date,
            COALESCE(`Go_ahead_Date`, received_date),
            COALESCE(`Reopen_dates`, received_date)
        ) AS start_date
    FROM T1
)

/*
 * Everything from here down (due_date_Recal, Ageing, Ageing_Bucket,
 * TAT Status_Recal, due_date_Final, opt_Final_due_date, Ageing_Final,
 * Ageing_Bucket_Final, TAT Status_Final, Insuff_Ageing,
 * Insuff_Ageing_Bucket) is computed in Python instead of SQL - see
 * tat_calculations.py. Those layers all depended on correlated per-row
 * subqueries against the holidays table, which this MySQL server executes
 * once per output row ("Range checked for each record") instead of once
 * per query - with ~10 such subqueries per row, that made even a
 * single-client, 10-day-window pull take minutes. Doing the same
 * calculation as one vectorized pandas pass after fetching (holidays
 * table is only ~3k rows, trivial to hold in memory) avoids that entirely.
 */
SELECT * FROM T2;

"""


def build_base_data_query(company_ids: Optional[Iterable[int]] = None, months_back: int = 4) -> str:
    """Build a ready-to-run QUERY_BASE_DATA, scoped to specific client(s) when given.

    company_ids: real client company_id values (from QUERY_CLIENT_LIST) to
        restrict the pull to - strongly recommended, since the unscoped
        "all clients" query is a multi-million-row pull on this database.
    months_back: how many months of received_date history to include.
    """
    if company_ids:
        ids = ",".join(str(int(cid)) for cid in company_ids)
        company_filter = f"AND emc.company_id IN ({ids})"
    else:
        company_filter = ""

    return _QUERY_BASE_DATA_TEMPLATE.format(months_back=int(months_back), company_filter=company_filter)
