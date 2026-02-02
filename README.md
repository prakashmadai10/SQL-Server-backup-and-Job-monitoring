---

## 📊 SQL Server Monitoring Automation

**Backup Status & Failed Job Reports**

This repository contains a Python-based automation that monitors **SQL Server Agent job failures** and **database backup status** across multiple SQL Server instances.
It generates **formatted Excel reports** and automatically **emails them to stakeholders**.

---

## 🚀 Features

### 💾 Backup Status Report (Last 24 Hours)

* Retrieves the **latest backup per database and backup type**:

  * Full
  * Differential
  * Log
* Shows:

  * Backup start & finish time
  * Backup size (MB & GB)
  * Backup destination path
  * Backup status (**SUCCESS / FAILED**)
* Highlights **failed or missing backups** in red
* Separate Excel sheets per SQL instance + consolidated view

### ⚠️ SQL Agent Jobs – Failed Only (Last 15 Days)

* Captures **only failed job steps**
* Includes:

  * Job name
  * Step ID & step name
  * Run date
  * Step duration (HH:MM:SS)
  * SQL severity & message
* Pulls failures from the **latest execution window**
* Red-highlighted rows for fast triage
* Per-instance sheets + combined report

### 📧 Automated Email Delivery

* Sends both reports as Excel attachments
* HTML-formatted email body
* Supports **To / CC / BCC**
* Ideal for daily DBA or operations monitoring

---

## 🗂 Project Structure

```
.
├── backup_report.py        # Backup status (last 24 hours)
├── jobs_report.py          # Failed SQL Agent jobs (last 15 days)
├── common.py               # Shared utilities (DB, Excel, email, logging)
├── run_all_and_email.py    # Orchestrator (runs reports + emails results)
├── .env                    # SQL credentials (not committed)
└── logs/
```

---

## ⚙️ Prerequisites

* Python **3.9+**
* SQL Server access to:

  * `msdb`
* ODBC Driver:

  * **ODBC Driver 17 for SQL Server**
* SMTP relay access (internal or external)

### Python Packages

```bash
pip install pandas pyodbc python-dotenv xlsxwriter
```

---

## 🔐 Environment Variables (`.env`)

Create a `.env` file in the project root:

```ini
SQL_USER=your_sql_user
SQL_PWD=your_sql_password
```

> If using Windows Authentication, set `USE_TRUSTED = True` in `run_all_and_email.py`.

---

## 🖥 Configuring SQL Instances

Edit the `INSTANCES` list in `run_all_and_email.py`:

```python
INSTANCES = [
    ("ProdDB", r"SERVER01\SQL2019"),
    ("Reporting", r"SERVER02\SQL2022"),
]
```

Each entry:

* **Label** → appears in Excel & email
* **Server name** → SQL Server instance

---

## ▶️ How to Run

```bash
python run_all_and_email.py
```

What happens:

1. Queries all configured SQL instances
2. Generates:

   * Backup status Excel file
   * Failed jobs Excel file
3. Applies conditional formatting
4. Emails reports to recipients

---

## 📈 Excel Output Details

### Backup Report

* File name:

  ```
  BackupStatus_Last24H_YYYY-MM-DD.xlsx
  ```
* Sheets:

  * `AllInstances`
  * One sheet per SQL instance
* Red rows = **FAILED backups**

### Jobs Report

* File name:

  ```
  JobSteps_Failed_15days_YYYY-MM-DD.xlsx
  ```
* Sheets:

  * `AllInstances`
  * One sheet per SQL instance
* Red rows = **Failed job steps**

---

## 🧠 Why This Exists

* Eliminates manual:

  * Job History checks
  * Backup validation
* Improves:

  * Early failure detection
  * Audit readiness
  * DBA operational efficiency
* Scales easily to **multiple SQL instances**

---

## 🛠 Future Enhancements (Optional Ideas)

* Slack / Teams notifications
* Threshold-based alerts (no failures → no email)
* Backup age SLA validation
* Job owner / schedule metadata
* Dockerized execution


