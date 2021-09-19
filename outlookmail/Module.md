## Overview
Ballerina connector for Microsoft Outlook mail provides access to the Microsoft Outlook mail service in Microsoft Graph v1.0 via the 
[Ballerina language](https://ballerina.io/). It provides the capability to perform more useful functionalities provided in Microsoft outlook mail such as sending messages, listing messages, creating drafts, mail folders, deleting messages, updating messages, etc.

This module supports [Microsoft Graph (Mail) API v1.0](https://docs.microsoft.com/en-us/graph/api/resources/message?view=graph-rest-1.0).

## Prerequisites
Before using this connector in your Ballerina application, complete the following:
* Create [Microsoft Outlook Account](https://outlook.live.com/owa/)
* Obtaining tokens
1. Follow [this link](https://docs.microsoft.com/en-us/graph/auth-v2-user#authentication-and-authorization-steps) and obtain the client ID, client secret, and refresh token.
 
## Quickstart

To use the Outlook mail connector in your Ballerina application, update the .bal file as follows:

### Step 1 - Import connector
Import the `ballerinax/microsoft.outlook.mail` module into the Ballerina project.
```ballerina
import ballerinax/microsoft.outlook.mail;
```
### Step 2: - Create a new connector instance
You can now make the connection configuration using the OAuth2 refresh token grant config.
```ballerina
mail:ConnectionConfig configuration = {
    auth: {
        refreshUrl: <REFRESH_URL>,
        refreshToken : <REFRESH_TOKEN>,
        clientId : <CLIENT_ID>,
        clientSecret : <CLIENT_SECRET>
    }
};

mail:Client outlookClient = check new(configuration);
```
### Step 3: Invoke connector operation
1. Send a new message using the sendMessage remote operation
```ballerina
public function main() returns error? {
    mail:MessageContent messageContent = {
        message: {
            subject: "Ballerina Test Email",
            importance: "Low",
            body: {
                "contentType": "HTML",
                "content": "This is sent by sendMessage operation <b>Test</b>!"
            },
            toRecipients: [
                {
                    emailAddress: {
                        address: "<email address>",
                        name: "<name>"
                    }
                }
            ]
        },
        saveToSentItems: true
    };
    http:Response response = check oneDriveClient->sendMessage(messageContent);
}
``` 
2. Use `bal run` command to compile and run the Ballerina program.

**[You can find more samples here](https://github.com/ballerina-platform/module-ballerinax-microsoft.outlook.mail/tree/main/outlookmail/samples)**
