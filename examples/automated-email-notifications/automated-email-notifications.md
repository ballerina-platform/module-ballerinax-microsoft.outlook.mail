# Automated email notifications with Microsoft Outlook Mail

## Introduction

This guide demonstrates how to automate weekly project status report distribution using the Microsoft Outlook Mail connector for Ballerina. The workflow covers creating a dedicated mail folder, drafting an HTML-formatted email, attaching a file to the draft, sending it to a recipient, and confirming delivery by listing recent messages.

## Prerequisites

Follow the guidelines in the [Setup guide](https://github.com/ballerina-platform/module-ballerinax-microsoft.outlook.mail#setup-guide) to obtain the necessary credentials to access the Microsoft Outlook Mail API.

> **Note:** This example creates a mail folder named `Weekly Status Reports`. If a folder with that name already exists in your account, either delete it beforehand or update the folder name in the code to avoid conflicts.

### Configuration

Configure the Microsoft Outlook Mail API credentials in `Config.toml` in the example directory.

```toml
refreshUrl = "<REFRESH_URL>"
refreshToken = "<REFRESH_TOKEN>"
clientId = "<CLIENT_ID>"
clientSecret = "<CLIENT_SECRET>"
recipientEmail = "<RECIPIENT_EMAIL>"
recipientName = "<RECIPIENT_NAME>"
```

## Run the example

Execute the following command to run the example.

```bash
bal run
```
