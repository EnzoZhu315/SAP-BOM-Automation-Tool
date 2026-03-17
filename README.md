Background & Business Impact
In manufacturing and supply chain management, maintaining the Bill of Materials (BOM) is a critical yet labor-intensive task. Manual entry of material components into SAP is prone to human error and significant time delays.

This project was developed to automate the CS01 (Create BOM) and CS02 (Change BOM) transactions. By bridging Google Sheets (as a cloud-based task manager) with SAP ERP, the solution ensures data integrity, real-time status tracking, and "zero-touch" processing for material maintenance.

Core Features
Cloud-Driven Task Queue: Automatically fetches pending "P-Numbers" (Material IDs) from a designated Google Sheet.

Bi-Directional Feedback: Upon successful execution in SAP, the script writes a "success" status back to the Google Sheet, providing a real-time audit trail.

Intelligent Dialog Handling: Automatically detects and bypasses SAP modal windows, "C9" status warnings, and date consistency alerts.

Credential Localization: Implements secure, temporary local handling of API JSON secrets to bypass network drive latency or access restrictions.

Technical Stack & Implementation
Python & Win32Com: Utilized Python's win32com library to interface with the SAP GUI Scripting Engine.

SAP Tracker/Recorder: Leveraged SAP's native recording tools to map complex GUI objects and ID paths for robust element targeting.

Google Sheets API (gspread): Integrated gspread and oauth2client for secure, OAuth2-authenticated cloud data access.

Exception Resilience: Built a "Retry & Enter" loop to handle SAP’s non-critical yellow warnings, ensuring batch processing is never interrupted.

System Architecture
Extraction: Script connects to Google Sheets API and filters for rows where Status != "success".

Execution:

CS01 Module: Handles header data, item categories (L), components, quantities, and units.

CS02 Module: Checks for existing components and appends/updates material lines using Change Numbers.

Verification: Scans the SAP Status Bar (SBar) for success keywords (e.g., "Changed", "Created").

Closing the Loop: Updates the specific row index in the cloud with a timestamped success marker.

Deployment (Portable Mode)
This project is designed to run via a Portable Python Environment, allowing deployment on corporate machines without administrative privileges or local Python installations.
