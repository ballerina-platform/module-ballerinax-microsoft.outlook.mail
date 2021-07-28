## Overview
Ballerina connector for Microsoft Outlook mail provides access to Microsoft Outlook mail service in Microsoft Graph v1.0 via Ballerina language easily. It provides capability to perform more useful functionalities provided in Microsoft outlook mail such as sending messages, listing messages, creating drafts, mail folders, deleting and updating messages etc. 

This module supports Microsoft Graph (Mail) API v1.0 version.

## Prerequisites
Before using this connector in your Ballerina application, complete the following:
* Create [Microsoft Outlook Account](https://outlook.live.com/owa/)
* Obtaining tokens
        
    Follow [this link](https://docs.microsoft.com/en-us/graph/auth-v2-user#authentication-and-authorization-steps) and obtain the client ID, client secret and refresh token.
 
## Quickstart

### For Choreo user

Choreo is a digital innovation platform that allows you to develop, deploy, and manage cloud-native applications at scale. Its AI-assisted, low-code application development environment simplifies creating services, managing APIs, and building integrations while ensuring best practices and secure coding guidelines

#### Step 1: Select Outlook.mail connector from the connector list
* Click the last **+** icon in the low-code diagram
* Click API Calls and then select the **Outlook.mail connector**
#### Step 2: Provide configuration detail
* Provide a name for the configuration
* Next, select the configuration type from the **OauthClientConfig** list
    1. For **BearerTokenConfig** type, Provide Bearer token in the next text area labeled as Token
    2. For **OAuth2RefreshTokenGrant** type, Provide client ID, client secret, refresh token and refresh URL as shown in the form labels
#### Step 3: Select an operation
* After providing the required details, Click on the **Continue To Invoke API** button and select an operation from the drop down list
* Provide input parameters for the operation as necessary

### For Ballerina user
To use the Outlook mail connector in your Ballerina application, update the .bal file as follows:

### Step 1: Import MS Outlook Mail Package
First, import the `ballerinax/microsoft.outlook.mail` module into the Ballerina project.
```ballerina
import ballerinax/microsoft.outlook.mail;
```
### Step 2: Configure the connection to an existing Azure AD app
You can now make the connection configuration using the OAuth2 refresh token grant config.
```ballerina
mail:Configuration configuration = {
    clientConfig: {
        refreshUrl: <REFRESH_URL>,
        refreshToken : <REFRESH_TOKEN>,
        clientId : <CLIENT_ID>,
        clientSecret : <CLIENT_SECRET>
    }
};

mail:Client outlookClient = check new(configuration);
```
### Step 3: Send a message
Send a new message using the sendMessage remote operation
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
## Quick reference 
The following code snippets shows how the connector operations can be used in different scenarios after initializing the client.
* Get detail of a message from a mailbox
 ```ballerina
   mail:Message message = check outlookClient->getMessage("<Message ID>");
   ```
* Send an existing draft
 ```ballerina
    _= check outlookClient->sendDraftMessage("<Draft ID>");
```
* List messages
```ballerina
    stream<mail:Message, error?> result = check outlookClient->listMessages("<Folder ID>", 
        optionalUriParameters = "?$select: \"sender,subject,hasAttachments\"");
```

***[You can find more samples here](https://github.com/ballerina-platform/module-ballerinax-microsoft.outlook.mail/tree/main/outlookmail/samples)***
