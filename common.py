import os
import math
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime
import mimetypes
import smtplib
from email.message import EmailMessage

import pyodbc
import pandas as pd
from dotenv import load_dotenv


def load_env(script_dir: str) -> dict:
    env_path = os.path.join(script_dir, ".env")
    if not os.path.exists(env_path):
        raise FileNotFoundError(f".env not found: {env_path}")
    load_dotenv(dotenv_path=env_path, override=True)

    return {
        "SQL_USER": os.getenv("SQL_USER"),
        "SQL_PWD": os.getenv("SQL_PWD"),
    }


def setup_logger(name: str, log_file: str) -> logging.Logger:
    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)

    if logger.handlers:
        return logger

    file_handler = RotatingFileHandler(
        log_file, maxBytes=1_000_000, backupCount=5, encoding="utf-8"
    )
    file_handler.setFormatter(logging.Formatter(
        "%(asctime)s [%(levelname)s] %(name)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    ))

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(logging.Formatter(
        "%(asctime)s [%(levelname)s] - %(message)s",
        datefmt="%H:%M:%S",
    ))

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    return logger


def build_conn_str(server: str, use_trusted: bool, sql_user: str | None, sql_pwd: str | None) -> str:
    if use_trusted:
        return (
            "DRIVER={ODBC Driver 17 for SQL Server};"
            f"SERVER={server};DATABASE=msdb;Trusted_Connection=yes;"
        )
    if not sql_user or not sql_pwd:
        raise EnvironmentError("SQL_USER/SQL_PWD missing while USE_TRUSTED=False")
    return (
        "DRIVER={ODBC Driver 17 for SQL Server};"
        f"SERVER={server};DATABASE=msdb;"
        f"UID={sql_user};PWD={sql_pwd};"
    )


def excel_col_letter(col_idx_zero_based: int) -> str:
    col = col_idx_zero_based + 1
    letters = ""
    while col:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def msdb_datetime(run_date):
    if run_date is None:
        return None

    d = str(int(run_date)).zfill(8)  # YYYYMMDD
    return datetime.strptime(d, "%Y%m%d").date()


def msdb_duration(run_duration):
    if run_duration is None:
        return None
    s = str(run_duration).zfill(6)  # HHMMSS
    hh = int(s[0:2])
    mm = int(s[2:4])
    ss = int(s[4:6])
    return f"{hh:02d}:{mm:02d}:{ss:02d}"


def bytes_to_mb(x):
    if x is None:
        return None
    if isinstance(x, float) and math.isnan(x):
        return None
    return round(int(x) / 1024 / 1024, 2)


def bytes_to_gb(x):
    if x is None:
        return None
    if isinstance(x, float) and math.isnan(x):
        return None
    return round(int(x) / 1024 / 1024 / 1024, 2)


def write_excel_with_instance_sheets(
    output_file: str,
    df_all: pd.DataFrame,
    instance_data: dict[str, pd.DataFrame],
    status_col: str,
    fail_value: str,
    wrap_cols: set[str] | None = None,
    datetime_format: str = "yyyy-mm-dd hh:mm:ss"
):
    wrap_cols = wrap_cols or set()

    with pd.ExcelWriter(output_file, engine="xlsxwriter", datetime_format=datetime_format) as writer:
        workbook = writer.book
        border_format = workbook.add_format({"border": 1})
        fail_format = workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006", "border": 1})
        wrap_format = workbook.add_format({"text_wrap": True, "border": 1})

        def write_sheet(name: str, df: pd.DataFrame):
            df.to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]

            # Auto-fit columns, wrap selected columns
            for col_idx, col in enumerate(df.columns):
                max_len = max([len(str(col))] + [len(str(x)) for x in df[col].astype(str).tolist()])
                width = min(max_len + 2, 100)

                if col in wrap_cols:
                    ws.set_column(col_idx, col_idx, max(width, 80), wrap_format)
                else:
                    ws.set_column(col_idx, col_idx, min(width, 60), border_format)

            # Header formatting (also bordered)
            ws.set_row(0, None, border_format)


            last_row = len(df) + 1
            last_col_letter = excel_col_letter(len(df.columns) - 1)

            ws.conditional_format(
                f"A1:{last_col_letter}{last_row}",
                {"type": "no_blanks", "format": border_format},
            )

            status_letter = excel_col_letter(df.columns.get_loc(status_col))
            ws.conditional_format(
                f"A2:{last_col_letter}{last_row}",
                {"type": "formula", "criteria": f'=${status_letter}2="{fail_value}"', "format": fail_format},
            )

        write_sheet("AllInstances", df_all)

        for label, df_inst in instance_data.items():
            safe = label[:31].replace("/", "_").replace("\\", "_").replace(":", "_")
            write_sheet(safe, df_inst)


def send_email_with_attachments(
    smtp_server: str,
    smtp_port: int,
    from_email: str,
    to_emails: list[str],
    cc_emails: list[str],
    bcc_emails: list[str],
    subject: str,
    html_body: str,
    attachments: list[str],
    logger: logging.Logger
):
    msg = EmailMessage()
    msg["From"] = from_email
    msg["To"] = ", ".join(to_emails)
    if cc_emails:
        msg["Cc"] = ", ".join(cc_emails)
    msg["Subject"] = subject
    msg.add_alternative(html_body, subtype="html")

    for path in attachments:
        if not os.path.exists(path):
            raise FileNotFoundError(f"Attachment not found: {path}")
        mime_type, _ = mimetypes.guess_type(path)
        maintype, subtype = (mime_type.split("/") if mime_type else ("application", "octet-stream"))
        with open(path, "rb") as f:
            msg.add_attachment(f.read(), maintype=maintype, subtype=subtype, filename=os.path.basename(path))

    recipients = to_emails + cc_emails + bcc_emails
    logger.info(f"Sending email to {recipients} via {smtp_server}:{smtp_port}")

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.send_message(msg, from_addr=from_email, to_addrs=recipients)

    logger.info("Email sent successfully")
