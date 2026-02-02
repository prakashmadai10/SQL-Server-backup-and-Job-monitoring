import os
from datetime import datetime

from common import load_env, setup_logger, send_email_with_attachments
from jobs_report import generate_jobs_report
from backup_report import generate_backup_report


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    creds = load_env(script_dir)

    logger = setup_logger("db_monitor", os.path.join(script_dir, "C:\DB_Instances_Monitor_Jobs\logs\db_monitor.log"))
    logger.info("===== Combined Monitor Started =====")

    # =====================
    # CONFIG (single place)
    # =====================
    INSTANCES = [
    # label, server_name
    ("Nickname", r"machine_name\instance"),  # add/remove as needed4
        ("Nickname", r"machine_name\instance2"),  # add/remove as needed



]

    USE_TRUSTED = False

    # --- SMTP / Email Settings ---
    SMTP_SERVER = "xxx."  # or valid IP like "18.13.2.58"
    SMTP_PORT = 25                 # usually 25 or as given by IT

    FROM_EMAIL = "xyz@gmail.com"
    TO_EMAILS = [
        "abcd@gmail.com"
    ]
    CC_EMAILS = [
        # "manager@yourcompany.com",
    ]
    BCC_EMAILS = [
        # "you@yourcompany.com",
    ]
    today = datetime.now().strftime("%Y-%m-%d")

    jobs_file = os.path.join(script_dir, f"C:\DB_Instances_Monitor_Jobs\csv_exported\JobSteps_Failed_15days_{today}.xlsx")
    backup_file = os.path.join(script_dir, f"C:\DB_Instances_Monitor_Jobs\csv_exported\BackupStatus_Last24H_{today}.xlsx")
    # Generate both reports
    jobs_path = generate_jobs_report(
        instances=INSTANCES,
        use_trusted=USE_TRUSTED,
        sql_user=creds["SQL_USER"],
        sql_pwd=creds["SQL_PWD"],
        output_file=jobs_file,
        logger=logger
        
    )

    backup_path = generate_backup_report(
        instances=INSTANCES,
        use_trusted=USE_TRUSTED,
        sql_user=creds["SQL_USER"],
        sql_pwd=creds["SQL_PWD"],
        output_file=backup_file,
        logger=logger
    )

    



    subject = f"DB Monitoring Report (Jobs + Backups) – {today}"
    html_body = f"""
<html>
  <body style="font-family: Arial, Helvetica, sans-serif; font-size: 13px; color:#000;">
    <p><b>Attached are today’s automated DB monitoring reports ({today}).</b></p>

    <p>
      This automation consolidates monitoring for <b>SQL Agent jobs</b> and <b>SQL Server backups</b>
      across all configured instances, helping the team detect issues early and reduce time spent
      checking Job History and backup logs manually.
    </p>

    <hr>

    <p><b>📄 Report 1: SQL Agent Jobs (Latest Run)</b></p>
    <ul>
      <li>Latest run for each job, including <b>job outcome (step 0)</b> and <b>all executed steps</b>.</li>
      <li>Key columns:
        <ul>
          <li><b>Instance</b>, <b>JobName</b> → where the job ran</li>
          <li><b>RunStatus</b> → Succeeded/Failed/Retry/Canceled</li>
          <li><b>StepDuration</b> → HH:MM:SS</li>
          <li><b>Message</b> → detailed SQL Agent message / error</li>
        </ul>
      </li>
    </ul>

    <p><b>💾 Report 2: Backup Status (Last 24 Hours)</b></p>
    <ul>
      <li>Latest backup per database and type (<b>Full/Diff/Log</b>) within the last 24 hours.</li>
      <li>Includes <b>backup size</b> (logical + compressed) and <b>backup path</b> for validation.</li>
      <li>Key columns: <b>backup_status</b>, <b>physical_device_name</b>, <b>compressed_size_gb</b>.</li>
    </ul>

    <p><b>🎨 Color Legend</b></p>
    <ul>
      <li><span style="color:#9C0006;"><b>Red rows</b></span> = requires review:
        <ul>
          <li>Job failed at step/outcome, or</li>
          <li>Backup failed/missing for the last 24 hours.</li>
        </ul>
      </li>
    </ul>

    <p><b>🛠 Recommended Actions (when you see red rows)</b></p>
    <ul>
      <li><b>Jobs:</b> open Job History for the job/step, review the Message column, validate connectivity/credentials/disk space, and rerun if appropriate.</li>
      <li><b>Backups:</b> confirm backup destination availability, check SQL Agent backup jobs, verify disk space, and rerun backups if needed.</li>
      <li>If escalation is needed, reply with <b>Instance + JobName</b> (jobs) or <b>Instance + Database + Backup Type</b> (backups).</li>
    </ul>

    <p><b>✅ Why this automation exists</b></p>
    <ul>
      <li>Early detection of failures to protect downstream systems and recovery readiness.</li>
      <li>Consistent monitoring across all instances.</li>
      <li>Reduced manual effort and improved audit readiness.</li>
    </ul>

  </body>
</html>
"""

    send_email_with_attachments(
        smtp_server=SMTP_SERVER,
        smtp_port=SMTP_PORT,
        from_email=FROM_EMAIL,
        to_emails=TO_EMAILS,
        cc_emails=CC_EMAILS,
        bcc_emails=BCC_EMAILS,
        subject=subject,
        html_body=html_body,
        attachments=[jobs_path, backup_path],
        logger=logger
    )

    logger.info("===== Combined Monitor Finished =====")


if __name__ == "__main__":
    main()
