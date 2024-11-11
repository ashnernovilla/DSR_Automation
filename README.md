# Daily Submissions Report Automation

This project automates the process of generating daily submissions reports. The goal is to reduce manual efforts in report generation and improve accuracy and timeliness in sharing submissions data. By running a scheduled process, the system automatically extracts, processes, and delivers the report to specified stakeholders.

## Features

- **Automated Report Generation**: Schedule daily data extraction and report generation.
- **Data Transformation**: Cleans and structures raw submission data for easy analysis.
- **Customizable Output Formats**: Generates reports in CSV, Excel, or PDF formats.
- **Email Notifications**: Sends the daily report to designated recipients.
- **Error Handling & Logging**: Captures errors and logs them for easy troubleshooting.
- **Configurable Settings**: Customize parameters like email recipients, report time, and data sources.

## Table of Contents

- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [Scheduled Automation](#scheduled-automation)
- [Conceptual Diagram](#conceptual-diagram)


## Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/ashnernovilla/daily-submissions-report-automation.git
   cd daily-submissions-report-automation.

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt


## Configuration

1. **Setup the Folders Location and Lifelines**

2. **Administrator Rights**

3. **Correct Python Evnironment**

## Usage

1. **Setup the Folders Location and Lifelines**:
   ```bash
   python DSR_Auto.py

## Scheduled Automation
To automate the report generation, use a scheduler like cron (Linux) or Task Scheduler (Windows) to run generate_report.py daily at the desired time.

1. Example (Windows Task Scheduler)
2. Open Task Scheduler and create a new task.
3. Set the "Trigger" to "Daily" at your preferred time.
4. Set the "Action" to run DSR_Auto.py with your Python interpreter.

## Conceptual Diagram
![image](https://github.com/user-attachments/assets/cce4cd03-0bee-4080-9363-c2411ab6df7c)

We welcome contributions to enhance functionality, fix bugs, or improve documentation.

Fork the project.
Create a feature branch.
Commit your changes and push the branch.
Open a Pull Request.


