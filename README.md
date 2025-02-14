# Data Download Automation
Overview
Data Download Automation is a Python-based solution designed to streamline the processing of data exporting from emails from Microsoft Outlook. The script automates the extraction, transformation, and distribution of review and catalog data. It connects to Outlook to retrieve specific emails, extracts file URLs from the email bodies, downloads and transforms Excel files using Pandas, and finally, sends the processed files via email—all while maintaining robust logging for transparency and debugging.

## Features
- Connects to Microsoft Outlook using win32com to fetch emails based on predefined subjects.
- Utilises regular expressions to extract download URLs directly from email content.
- Automatically opens the URL in a web browser, constructs file paths, and waits for files to appear in the download folder before processing.
- Processes and cleans review and catalog Excel files with Pandas. This includes date transformations, column reordering, and data cleaning.
- Composes and sends an email with the transformed files attached using Outlook’s COM interface.
- Implements detailed logging to track each step of the process, making debugging and monitoring straightforward.

## Technologies Used
### Python
- win32com: For interfacing with Microsoft Outlook
- Pandas: For data manipulation and transformation
- Regular Expressions (re): For parsing and extracting URLs
- Threading: For file monitoring and wait handling
- Logging: For process tracking and debugging

## Roadmap
- Externalise and secure configuration parameters (e.g., email subjects, folder paths) for easier maintenance and scalability.
- Integrate with task schedulers (e.g., Azure Scheduler) to execute the script at regular intervals without manual intervention.
