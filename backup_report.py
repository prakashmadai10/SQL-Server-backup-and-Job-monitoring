import pandas as pd
import pyodbc

from common import (
    build_conn_str, bytes_to_mb, bytes_to_gb,
    write_excel_with_instance_sheets
)

SQL_BACKUP = """
WITH LatestBackups AS (
    SELECT
        bs.database_name,
        bs.type AS backup_type_code,
        CASE bs.type
            WHEN 'D' THEN 'Full'
            WHEN 'I' THEN 'Differential'
            WHEN 'L' THEN 'Log'
        END AS backup_type,
        bs.backup_start_date,
        bs.backup_finish_date,
        bs.backup_size,
        bmf.physical_device_name,
        ROW_NUMBER() OVER (
            PARTITION BY bs.database_name, bs.type
            ORDER BY bs.backup_finish_date DESC
        ) AS rn
    FROM msdb.dbo.backupset bs
    JOIN msdb.dbo.backupmediafamily bmf
        ON bs.media_set_id = bmf.media_set_id
    WHERE bs.backup_start_date >= DATEADD(HOUR, -24, GETDATE())
)
SELECT
    database_name,
    backup_type,
    backup_start_date,
    backup_finish_date,
    backup_size,
    physical_device_name,
    CASE
        WHEN backup_finish_date IS NULL THEN 'FAILED'
        ELSE 'SUCCESS'
    END AS backup_status
FROM LatestBackups
WHERE rn = 1
ORDER BY database_name, backup_type;
"""

def generate_backup_report(
    instances, use_trusted, sql_user, sql_pwd, output_file, logger
) -> str:
    instance_data = {}
    all_frames = []

    for label, server in instances:
        logger.info(f"[BACKUP] Querying instance '{label}' ({server})")
        conn_str = build_conn_str(server, use_trusted, sql_user, sql_pwd)

        try:
            with pyodbc.connect(conn_str) as conn:
                df = pd.read_sql(SQL_BACKUP, conn)
        except Exception as e:
            logger.exception(f"[BACKUP] Failed query {label} ({server}): {e}")
            continue

        if df.empty:
            logger.warning(f"[BACKUP] No rows for {label}")
            continue

        df["Instance"] = label
        df["backup_size_mb"] = df["backup_size"].apply(bytes_to_mb)
        df["backup_size_gb"] = df["backup_size"].apply(bytes_to_gb)
      

        df = df[
            [
                "Instance",
                "database_name",
                "backup_type",
                "backup_start_date",
                "backup_finish_date",
                "backup_status",
                "backup_size_mb",
                "backup_size_gb",
                "physical_device_name",
            ]
        ]

        instance_data[label] = df
        all_frames.append(df)

    if not all_frames:
        raise RuntimeError("[BACKUP] No data collected from any instance.")

    df_all = pd.concat(all_frames, ignore_index=True)

    write_excel_with_instance_sheets(
        output_file=output_file,
        df_all=df_all,
        instance_data=instance_data,
        status_col="backup_status",
        fail_value="FAILED",
        wrap_cols={"physical_device_name"}
    )

    logger.info(f"[BACKUP] Excel created: {output_file}")
    return output_file
