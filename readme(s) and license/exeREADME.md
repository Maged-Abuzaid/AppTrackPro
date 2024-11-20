# AppTrackPro

![AppTrackPro Logo](assets/app_icon.ico)

## Introduction

**AppTrackPro** is a user-friendly application designed to help you manage and track your job applications efficiently. With a sleek interface and powerful features, AppTrackPro ensures you stay organized and never miss an opportunity.

## Features

- **Add Applications:** Easily input details of your job applications.
- **View & Edit:** Manage your applications with options to search, edit, and delete entries.
- **Google Sync:** Seamlessly synchronize your application data with Google Sheets.
- **Personal Info Management:** Store and access your personal information securely.
- **Theming:** Choose between Dark and Light modes to suit your preference.
- **Drag-and-Drop:** Effortlessly upload essential files using drag-and-drop functionality.

## Getting Started

### Prerequisites

- **Operating System:** Windows 10 or later
- **Python:** Not required (executables are bundled)
- **Internet Connection:** Required for Google (not local save)

### Downloading the Executable

1. **Download the .exe File**
   - Visit the [AppTrackPro Releases](https://github.com/Maged-Abuzaid/AppTrackPro/releases) page on GitHub.
   - Download the latest `AppTrackPro.exe` file.

2. **Run the Application**
   - Double-click the downloaded `AppTrackPro.exe` file.
   - The application will launch, displaying the main window.

### Setting Up and Configuring Google Sync

To enable Google Sync and synchronize your application data with Google Sheets, follow these steps:

#### 1. Setting Up a Google API

1. **Create a Google Cloud Project**
   - Go to the [Google Cloud Console](https://console.cloud.google.com/).
   - Click on **Select a project** and then **New Project**.
   - Enter a project name and click **Create**.

2. **Enable Google Sheets API**
   - In the Google Cloud Console, navigate to **APIs & Services > Library**.
   - Search for "Google Sheets API" and click **Enable**.

3. **Create Service Account Credentials**
   - Go to **APIs & Services > Credentials**.
   - Click on **Create Credentials > Service Account**.
   - Provide a name and description for the service account.
   - Click **Create and Continue**, then **Done**.
   - Once the service account is created, click on it to open its details.
   - Navigate to the **Keys** tab and click **Add Key > Create New Key**.
   - Select **JSON** and click **Create** to download the `service_account.json` file.

4. **Share Your Google Sheet with the Service Account**
   - Open your Google Sheet or create a new one.
   - Click on **Share** and enter the service account's email address (found in the `service_account.json` file).
   - Assign **Editor** permissions and click **Send**.

#### 2. Uploading Required Files

1. **Service Account JSON File**
   - In AppTrackPro, navigate to **Settings** (⚙️ icon in the menu bar).
   - Click on **Google Sync Configuration**.
   - Upload the `service_account.json` file you downloaded earlier.

2. **Google Sheets ID**
   - Open your Google Sheet and copy the Spreadsheet ID from the URL.
     - Example URL: `https://docs.google.com/spreadsheets/d/your_spreadsheet_id/edit#gid=0`
     - The part between `/d/` and `/edit` is your Spreadsheet ID.
   - In AppTrackPro's **Google Sync Configuration**, enter the Spreadsheet ID.

#### 3. (Optional) Uploading an Existing Applications.xlsx

If you have an existing `Applications.xlsx` file with the following columns:
- **Company**
- **Position**
- **Application Portal URL**
- **Date Applied**
- **Status**

You can upload it to AppTrackPro:

1. **Navigate to Settings**
   - Click on the **Settings** (⚙️) icon in the menu bar.
   - Go to **Google Sync Configuration**.

2. **Upload Applications.xlsx**
   - Click on the designated area to upload the `Applications.xlsx` file.
   - Alternatively, drag and drop the file into the upload zone.

### Using AppTrackPro

1. **Adding a New Application**
   - Open AppTrackPro.
   - Navigate to the **Add Application** tab.
   - Fill in the **Company**, **Position**, and **Application Portal URL** fields.
   - Click **Submit** to save the application.

2. **Viewing and Editing Applications**
   - Switch to the **View/Edit Applications** tab.
   - Use the search bar to filter applications.
   - Right-click on any application to delete or copy it.
   - Double-click on cells to edit their contents directly.

3. **Managing Personal Information**
   - Go to the **Clipboard** tab.
   - Edit your personal information as needed.
   - Click **Save** to store changes.

4. **Switching Themes**
   - Click on the **Settings** (⚙️) icon in the menu bar.
   - Select **Switch Theme** to toggle between Dark and Light modes.

### Troubleshooting

- **Google Sync Issues:**
  - Ensure that the `service_account.json` file is correctly uploaded.
  - Verify that the Spreadsheet ID is accurate.
  - Make sure the service account has Editor access to the Google Sheet.
  - Check your internet connection.

- **Application Not Launching:**
  - Ensure that you have downloaded the correct executable for your operating system.
  - Disable any antivirus software that might be blocking the application.
  - Re-download the executable in case the file was corrupted during download.

- **Data Not Syncing:**
  - Confirm that Google Sync is enabled in Settings.
  - Check if there are any errors displayed in the application logs.

### Support

For further assistance, please contact [MagedM.Abuzaid@gmail.com](mailto:MagedM.Abuzaid@gmail.com) or visit the [AppTrackPro GitHub Issues](https://github.com/Maged-Abuzaid/AppTrackPro/issues) page.

---