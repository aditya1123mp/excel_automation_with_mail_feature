Here's a sample README for your GitHub repository that describes the automation for the Orange HRM site:

---

# Orange HRM Automation Suite

This repository contains automation scripts for the Orange HRM site, focusing on the "Add Employee" functionality. The automation workflow involves creating new employees, saving the employee data to an Excel file, sending that file via email, downloading the file to a specific folder, and verifying the created employees.

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Setup](#setup)
- [Usage](#usage)
  - [Add Employee](#add-employee)
  - [Send Excel via Email](#send-excel-via-email)
  - [Download Excel Attachment](#download-excel-attachment)
  - [Verify Employees](#verify-employees)
- [Contributing](#contributing)
- [License](#license)

## Overview

This project automates the "Add Employee" functionality of the Orange HRM application. The automation process is divided into multiple steps:

1. **Add Employee**: Automates the process of adding a new employee to the system.
2. **Save Employee Data**: Saves the details of the newly created employees in an Excel file.
3. **Send Excel via Email**: Sends the generated Excel file containing employee data to a specified email address.
4. **Download Excel Attachment**: A separate script downloads the Excel file attachment from the email and saves it in a designated folder.
5. **Verify Employees**: The final script reads the downloaded Excel sheet and verifies that all the employees listed have been successfully created in the Orange HRM application.

## Features

- Automated "Add Employee" functionality.
- Employee data is saved in an Excel file.
- The generated Excel file is automatically sent via email.
- Script to download the Excel file attachment from the email.
- Script to read the Excel file and verify employee creation in the system.

## Prerequisites

- **Katalon Studio**: This project uses Katalon Studio for automation. Ensure you have the latest version installed.
- **Java**: Required for running Katalon Studio and some of the scripts.
- **Email Account**: An email account configured to send and receive emails.
- **Excel**: Microsoft Excel or any compatible software to view and edit the Excel file.

## Setup

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/orange-hrm-automation.git
   cd orange-hrm-automation
   ```

2. Open the project in Katalon Studio.

3. Configure your email credentials in the `emailConfig.properties` file.

4. Ensure that all dependencies are installed and configured correctly.

## Usage

### Add Employee

This script automates the process of adding a new employee to the Orange HRM system.

- Run the `AddEmployee` test case.
- The script will add a new employee and save their details in an Excel file.

### Send Excel via Email

This script sends the generated Excel file to a specified email address.

- Run the `SendExcelByEmail` test case.
- The script will attach the Excel file and send it to the configured email address.

### Download Excel Attachment

This script downloads the Excel file attachment from the email and saves it in a specific folder.

- Run the `DownloadExcelAttachment` test case.
- The script will download the email attachment and save it to the `DownloadedFiles` folder.

### Verify Employees

This script reads the downloaded Excel file and verifies that all the employees listed have been successfully created in the Orange HRM system.

- Run the `VerifyEmployees` test case.
- The script will check the employee details against the records in the Orange HRM system and report any discrepancies.

## Contributing

Contributions are welcome! Please fork this repository and submit a pull request with your changes.
