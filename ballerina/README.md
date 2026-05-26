## Overview

[Microsoft Outlook Mail](https://outlook.live.com/owa/) is a widely used email service from Microsoft, available as part of the Microsoft 365 suite.

The `ballerinax/microsoft.outlook.mail` connector offers APIs to connect and interact with the [Microsoft Outlook Mail API](https://learn.microsoft.com/en-us/graph/api/resources/mail-api-overview?view=graph-rest-1.0) endpoints, specifically based on the [Microsoft Graph REST API v1.0](https://learn.microsoft.com/en-us/graph/overview). It supports sending, receiving, and managing email messages, creating and organizing mail folders, managing file and item attachments, drafting and deleting messages, and reading or updating message properties such as subject, body, flags, and categories.

## Setup guide

To use the Microsoft Outlook Mail connector, you need a Microsoft account and an application registered in Azure Active Directory (Azure AD) with the appropriate OAuth2 credentials.

### Step 1: Sign in to Azure Portal

If you don't have a Microsoft Azure account, you can create one for free at [https://azure.microsoft.com](https://azure.microsoft.com).

Go to the [Azure Portal](https://portal.azure.com) and sign in with your Microsoft account. From the home page, navigate to **Entra**.

![Azure Portal](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.outlook.mail/refs/heads/main/docs/resources/azure-portal.png)

### Step 2: Register an application

1. In the Microsoft Entra admin center, navigate to **App registrations** from the left sidebar.

   ![Entra main page](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.outlook.mail/refs/heads/main/docs/resources/entra-main-page.png)

2. Click **New registration** in the top menu.

3. Fill in the application details.
   - **Name**: Provide a name for your app (e.g., `Ballerina Outlook Connector App`)
   - **Supported account types**: Select **Any Entra ID Tenant + Personal Microsoft Accounts**.
   - **Redirect URI**: Select **Web** and enter your redirect URI (e.g., `http://localhost` for local testing).

   ![Register application](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.outlook.mail/refs/heads/main/docs/resources/register-application.png)

4. Click **Register**.

### Step 3: Add API permissions

1. In your registered application, navigate to **API permissions** from the left sidebar.

2. Click **Add a permission** > **Microsoft Graph** > **Delegated permissions**.

   ![Add API permissions](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.outlook.mail/refs/heads/main/docs/resources/add-api-permissions.png)

3. Add the following permissions.
   - `Mail.Read`
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `MailboxSettings.Read`
   - `MailboxSettings.ReadWrite`
   - `offline_access`

4. Click **Add permissions**.

5. When prompted to grant consent, click **Accept** to allow the app access to the requested permissions.

   ![Allow permission](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.outlook.mail/refs/heads/main/docs/resources/allow-permission.png)

### Step 4: Get the client ID and client secret

1. Navigate to **Overview** in your registered application. Copy the **Application (client) ID** and save it as your `clientId`.

2. Navigate to **Certificates & secrets** > **Client secrets** > **New client secret**.

   ![Add client secret](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.outlook.mail/refs/heads/main/docs/resources/add-client-secret.png)

3. Provide a description, choose an expiry duration, and click **Add**.

4. Copy the generated **Value** of the secret and save it as your `clientSecret`.

### Step 5: Set up the authentication flow

Before using the connector, obtain a refresh token using the following OAuth2 authorization code flow:

1. Construct an authorization URL using the format below. Replace `<CLIENT_ID>`, `<REDIRECT_URI>` and `<SCOPE>` with your specific values:

   ```text
   https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=<CLIENT_ID>&response_type=code&redirect_uri=<REDIRECT_URI>&scope=<SCOPE>&response_mode=query
   ```

   Example values for `<SCOPE>`: `Mail.Read Mail.ReadWrite Mail.Send MailboxSettings.Read offline_access`

2. Open the URL in a browser and sign in with your Microsoft account. Grant the requested permissions.

3. After authorization, you will be redirected to your redirect URI with a `code` parameter in the URL. Copy this code.

4. Exchange the authorization code for tokens by running the following `curl` command. Replace the placeholder values with your specific values. For `SCOPE` you can use this `Mail.Read Mail.ReadWrite Mail.Send User.Read offline_access`

   ```bash
    curl --location 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token' \
    --header 'Content-Type: application/x-www-form-urlencoded' \
    --data-urlencode 'grant_type=authorization_code' \
    --data-urlencode 'code=<CODE>' \
    --data-urlencode 'redirect_uri=<REDIRECT_URI>' \
    --data-urlencode 'client_id=<CLIENT_ID>' \
    --data-urlencode 'client_secret=<CLIENT_SECRET>' \
    --data-urlencode 'scope=<SCOPE>'
   ```

   The response will contain your access token and refresh token.

   ```json
    {
        "token_type": "Bearer",
        "scope": "<SCOPE>",
        "refresh_token": "<REFRESH_TOKEN>",
        "access_token": "<ACCESS_TOKEN>",
        "expires_in": 3600
    }
   ```

5. Store the `refresh_token` securely for use in your application.

## Quickstart

To use the `microsoft.outlook.mail` connector in your Ballerina application, update your `.bal` file as follows:

### Step 1: Import the module

Import the `microsoft.outlook.mail` module.

```ballerina
import ballerinax/microsoft.outlook.mail;
```

### Step 2: Instantiate a new connector

1. Create a `Config.toml` file and configure the credentials obtained above:

   ```toml
   clientId = "<CLIENT_ID>"
   clientSecret = "<CLIENT_SECRET>"
   refreshToken = "<REFRESH_TOKEN>"
   refreshUrl = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"
   ```

2. Instantiate a `mail:ConnectionConfig` with the obtained credentials and initialize the connector with it.

   ```ballerina
   configurable string clientId = ?;
   configurable string clientSecret = ?;
   configurable string refreshToken = ?;
   configurable string refreshUrl = ?;

   final mail:Client outlookClient = check new ({
      auth: {
         clientId,
         clientSecret,
         refreshToken,
         refreshUrl
      }
   });
   ```

### Step 3: Invoke the connector operation

Now, utilize the available connector operations. A sample use case is shown below.

#### List the most recent messages in the mailbox

```ballerina
public function main() returns error? {
    mail:MicrosoftGraphMessageCollectionResponse response = check outlookClient->/me/messages.get(
        queries = {
            dollarTop: 5,
            dollarSelect: ["id", "subject", "from", "receivedDateTime", "isRead"]
        }
    );
    mail:MicrosoftGraphMessage[] messages = response.value ?: [];
    foreach mail:MicrosoftGraphMessage message in messages {
        io:println("Subject: ", message?.subject, " | Read: ", message?.isRead);
    }
}
```

## Examples

The `microsoft.outlook.mail` connector provides practical examples illustrating usage in various scenarios. Explore these [examples](https://github.com/ballerina-platform/module-ballerinax-microsoft.outlook.mail/tree/main/examples), covering the following use cases.

1. [Automated email notifications](https://github.com/ballerina-platform/module-ballerinax-microsoft.outlook.mail/tree/main/examples/automated-email-notifications): Automates weekly project status report distribution. Creates a dedicated mail folder, drafts an HTML-formatted report email with an attachment, sends the draft, and lists recent sent messages to confirm delivery.

2. [Email inbox management](https://github.com/ballerina-platform/module-ballerinax-microsoft.outlook.mail/tree/main/examples/email-inbox-management): Implements a customer support inbox triage workflow. Lists unread messages, fetches message details, marks messages as read after review, creates an organized folder for processed tickets, and deletes spam or resolved messages.
