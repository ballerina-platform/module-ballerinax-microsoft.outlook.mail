# Examples

The `microsoft.outlook.mail` connector provides practical examples illustrating usage in various scenarios. Explore these [examples](https://github.com/ballerina-platform/module-ballerinax-microsoft.outlook.mail/tree/main/examples).

1. [Automated email notifications](https://github.com/ballerina-platform/module-ballerinax-microsoft.outlook.mail/tree/main/examples/automated-email-notifications)

   This example shows how to automate weekly project status report distribution. It creates a dedicated mail folder, drafts an HTML-formatted report email, attaches a summary file, sends the draft, and lists recent sent messages to confirm delivery.

2. [Email inbox management](https://github.com/ballerina-platform/module-ballerinax-microsoft.outlook.mail/tree/main/examples/email-inbox-management)

   This example shows how to implement a customer support inbox triage workflow. It lists unread messages, fetches message details, marks messages as read after review, creates an organized folder for processed tickets, and deletes spam or resolved messages.

## Prerequisites

1. Follow the [instructions](https://github.com/ballerina-platform/module-ballerinax-microsoft.outlook.mail#set-up-guide) to set up the Microsoft Outlook Mail API.

2. For each example, create a `Config.toml` file in the example directory with your OAuth2 credentials. Here is an example of how your `Config.toml` file should look:

   ```toml
   clientId = "<CLIENT_ID>"
   clientSecret = "<CLIENT_SECRET>"
   refreshToken = "<REFRESH_TOKEN>"
   refreshUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
   ```

   The `automated-email-notifications` example also requires:

   ```toml
   recipientEmail = "<RECIPIENT_EMAIL>"
   recipientName = "<RECIPIENT_NAME>"
   ```

## Running an Example

Execute the following commands to build an example from the source:

* To build an example:

  ```bash
  bal build
  ```

* To run an example:

  ```bash
  bal run
  ```

## Building the Examples with the Local Module

**Warning**: Due to the absence of support for reading local repositories for single Ballerina files, the Bala of the module is manually written to the central repository as a workaround. Consequently, the bash script may modify your local Ballerina repositories.

Execute the following commands to build all the examples against the changes you have made to the module locally:

* To build all the examples:

  ```bash
  ./build.sh build
  ```

* To run all the examples:

  ```bash
  ./build.sh run
  ```
