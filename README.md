# Xj Ghack

# Assign Ease: Automated Employee Assignment and Project Management

Welcome to Assign Ease, a comprehensive solution for optimizing employee assignments and project management. This project leverages AppSheet, Google Sheets, Google Calendar, Gmail, and Google Apps Script to automate and streamline the process.

## Overview

Assign Ease provides a user-friendly interface for managing employees and projects. It automates the assignment of employees based on their skills and availability, integrates project events into their Google Calendars, and sends notifications through Gmail.

## Features

- **User Interaction via AppSheet:**
  - Manage employees and projects through an intuitive no-code platform.
  - Simplifies data entry and updates.

- **Data Management with Google Sheets:**
  - Centralized storage of employee and project data.
  - Seamless read/write operations between AppSheet and Google Sheets.

- **Automation with Apps Script:**
  - **Employee Assignment:** Automatically matches project requirements with available employees.
  - **Calendar Integration:** Adds project events to assigned employees' Google Calendars.
  - **Email Notifications:** Sends email notifications to assigned employees and alerts the project owner if assignments are insufficient.

## Technology Stack

- **AppSheet:** Provides the no-code user interface.
- **Google Sheets:** Central database for storing employee and project information.
- **Google Calendar:** Used for checking employee availability and scheduling project events.
- **Gmail:** Sends automated email notifications.
- **Google Apps Script:** Powers the automation logic for employee assignment, calendar integration, and email notifications.

## Getting Started

### Access Links

- **AppSheet Application:** [Assign Ease](https://www.appsheet.com/start/fb525c84-50d7-4550-a600-29412872fc48)
- **Google Sheet:** [Employee and Project Data](https://docs.google.com/spreadsheets/d/1m30iuAC4l4sbWMW9UBX_VTdNvZH5Vt_TvABaCUs6irw/edit?usp=sharing)
- **Apps Script:** [Assign Ease](https://script.google.com/d/1A6LbAX4ggzXsS6QmyXdwAl-TG4RBqiZvVg7nGncLoGXNBjMGZtfCKRgp/edit?usp=sharing)

### How It Works

1. **Add and Update Employees:**
   - Use the AppSheet application to enter and update employee information and skills.
   
2. **Create New Projects:**
   - Enter project details including required skills and the number of employees needed.

3. **Automated Assignment:**
   - When a new project is added, the Apps Script function is triggered to:
     - Match project requirements with available employees.
     - Check employee availability using their Google Calendar IDs.
     - Assign available employees to the project and update the Google Sheet.
     - Add project events to the assigned employees' Google Calendars.
     - Send email notifications to assigned employees.
     - Notify the project owner if there are insufficient assignments.

---

Thank you for checking out Assign Ease! For any questions or support, please contact [panzhinhuey@gmail.com](mailto:panzhinhuey@gmail.com) or [lawmeijun204@gmail.com](mailto:lawmeijun204@gmail.com).
