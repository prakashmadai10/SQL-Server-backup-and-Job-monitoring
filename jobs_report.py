import pandas as pd
import pyodbc

from common import (
    build_conn_str,  msdb_datetime, msdb_duration,
    write_excel_with_instance_sheets
)

SQL_JOBS = """
;WITH Outcomes AS (
    SELECT
        j.job_id,
        j.name AS JobName,
        h.instance_id,
        h.run_date,
        h.run_time,
        ROW_NUMBER() OVER (
            PARTITION BY j.job_id
            ORDER BY h.instance_id DESC
        ) AS rn
    FROM msdb.dbo.sysjobs j
    JOIN msdb.dbo.sysjobhistory h
      ON h.job_id = j.job_id
     AND h.step_id = 0
),
LatestOutcome AS (
    SELECT job_id, JobName, instance_id AS last_outcome_instance_id
    FROM Outcomes
    WHERE rn = 1
),
PrevOutcome AS (
    SELECT job_id, instance_id AS prev_outcome_instance_id
    FROM Outcomes
    WHERE rn = 2
)
SELECT
    lo.JobName,
    h.step_id,
    ISNULL(h.step_name, CASE WHEN h.step_id = 0 THEN '(Job Outcome)' ELSE '(No Step Name)' END) AS step_name,
    h.run_date,
    h.run_time,
    h.run_duration,
    CASE h.run_status
        WHEN 0 THEN 'Failed'
        WHEN 1 THEN 'Succeeded'
        WHEN 2 THEN 'Retry'
        WHEN 3 THEN 'Canceled'
        WHEN 4 THEN 'In Progress'
        ELSE 'Unknown'
    END AS RunStatus,
    h.sql_severity,
    h.sql_message_id,
    h.message,
    h.instance_id
FROM LatestOutcome lo
LEFT JOIN PrevOutcome po
  ON po.job_id = lo.job_id
JOIN msdb.dbo.sysjobhistory h
  ON h.job_id = lo.job_id
 AND h.instance_id <= lo.last_outcome_instance_id
 AND h.instance_id > ISNULL(po.prev_outcome_instance_id, 0)
 and  h.run_status = 0 
 and h.run_date >= CONVERT(
        int,
        CONVERT(char(8), DATEADD(DAY, -15, GETDATE()), 112)
  )
ORDER BY lo.JobName, h.step_id, h.instance_id;
"""

def generate_jobs_report(
    instances, use_trusted, sql_user, sql_pwd, output_file, logger,
    
) -> str:
    instance_data = {}
    all_frames = []
    for label, server in instances:
        logger.info(f"Querying instance '{label}' ({server})...")
        conn_str = build_conn_str(server, use_trusted, sql_user, sql_pwd)

        try:
            with pyodbc.connect(conn_str) as conn:
                df = pd.read_sql(SQL_JOBS, conn)
        except Exception as e:
            logger.exception(f"[JOBS] Failed query {label} ({server}): {e}")
            continue

        if df.empty:
            logger.warning(f"No job history returned for instance {label} ({server}).")
        else:
            logger.info(f"Retrieved {len(df)} rows from instance {label}.")

        # Convert to human-friendly fields
        df["RunDate"] = df["run_date"].apply(msdb_datetime)

        df["StepDuration"] = df["run_time"].apply(msdb_duration)
        df["Instance"] = label

        # Keep only clean export columns (no raw run_date/time/duration)
        df = df[
            [
                "Instance",
                "JobName",
                "step_id",
                "step_name",
                "RunDate",
                "StepDuration",
                "RunStatus",
                "sql_severity",
                "sql_message_id",
                "message",
            ]
        ]

        instance_data[label] = df
        all_frames.append(df)

    if not all_frames:
        raise RuntimeError("[JOBS] No data collected from any instance.")

    df_all = pd.concat(all_frames, ignore_index=True)

    write_excel_with_instance_sheets(
        output_file=output_file,
        df_all=df_all,
        instance_data=instance_data,
        status_col="RunStatus",
        fail_value="Failed",
        wrap_cols={"message"}
    )

    logger.info(f"[JOBS] Excel created: {output_file}")
    return output_file

