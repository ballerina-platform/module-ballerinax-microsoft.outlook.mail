## Overview
Ballerina connector for Microsoft Outlook mail provides access to Microsoft Outlook mail service in Microsoft Graph v1.0 via Ballerina language easily. It provides capability to perform more useful functionalities provided in Microsoft outlook mail such as sending messages, listing messages, creating drafts, mail folders, deleting and updating messages etc. 

This module supports Microsoft Graph (Mail) API v1.0 version.

## Prerequisites
Before using this connector in your Ballerina application, complete the following:
* Create [Microsoft Outlook Account](https://outlook.live.com/owa/)
* Obtaining tokens
        
    Follow [this link](https://docs.microsoft.com/en-us/graph/auth-v2-user#authentication-and-authorization-steps) and obtain the client ID, client secret and refresh token.
* Configure the connector with obtained tokens
 
## Quickstart

To use the Outlook mail connector in your Ballerina application, update the .bal file as follows:

Step 1: Import MS Outlook Mail Package
First, import the ballerinax/microsoft.outlook.mail module into the Ballerina project.
```ballerina
import ballerinax/microsoft.outlook.mail;
```
Step 2: Configure the connection to an existing Azure AD app
You can now make the connection configuration using the OAuth2 refresh token grant config.
```ballerina
outlookMail:Configuration configuration = {
    clientConfig: {
        refreshUrl: <REFRESH_URL>,
        refreshToken : <REFRESH_TOKEN>,
        clientId : <CLIENT_ID>,
        clientSecret : <CLIENT_SECRET>
    }
};

mail:Client outlookClient = check new(configuration);

```
Step 3: Send a message
```
public function main() returns error? {
    mail:DraftMessage draft = {
        subject:"<Mail Subject>",
        importance:"Low",
        body:{
            "contentType": "HTML",
            "content": "We are <b>Wso2</b>!"
        },
        toRecipients:[
            {
                emailAddress:{
                    address: "<Your Email Address>",
                    name: "Name (Optional)"
                }
            }
        ]
    };
    mail:Message message = check outlookClient->createMessage(draft);
    log:printInfo(message.toString());
}

``` 
## Quick reference 
* Get detail of a message from a mailbox
 ```ballerina
   mail:Message message = check outlookClient->getMessage("<Message ID>");
   ```
* Send an existing draft
 ```ballerina
    _= check outlookClient->sendDraftMessage("<Draft ID>");
```

### [You can find more samples here](https://github.com/ballerina-platform/module-ballerinax-microsoft.outlook.mail/tree/main/outlookmail/samples)