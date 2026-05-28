# Email inbox management with Microsoft Outlook Mail

## Introduction

This guide demonstrates a customer support inbox triage workflow using the Microsoft Outlook Mail connector for Ballerina. The workflow covers listing unread messages, fetching full message details, marking messages as read, creating a folder to organize processed tickets, retrieving folder details, and deleting resolved or spam messages.

## Prerequisites

Follow the guidelines in the [Setup guide](https://github.com/ballerina-platform/module-ballerinax-microsoft.outlook.mail#setup-guide) to obtain the necessary credentials to access the Microsoft Outlook Mail API.

> **Note:** This example creates a mail folder named `Customer Support`. If a folder with that name already exists in your account, either delete it beforehand or update the folder name in the code to avoid conflicts.

### Configuration

Configure the Microsoft Outlook Mail API credentials in `Config.toml` in the example directory.

```toml
refreshUrl = "<REFRESH_URL>"
refreshToken = "<REFRESH_TOKEN>"
clientId = "<CLIENT_ID>"
clientSecret = "<CLIENT_SECRET>"
```

## Run the example

Execute the following command to run the example.

```bash
bal run
```
