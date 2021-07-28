## Overview
Ballerina connector for Microsoft Outlook mail provides access to Microsoft Outlook mail service in Microsoft Graph v1.0 via Ballerina language easily. It provides capability to perform more useful functionalities provided in Microsoft outlook mail such as sending messages, listing messages, creating drafts, mail folders, deleting and updating messages etc. 

This module supports Microsoft Graph (Mail) API v1.0 version.

## Prerequisites
Before using this connector in your Ballerina application, complete the following:
1. Create [Microsoft Outlook Account](https://outlook.live.com/owa/)
2. Obtaining tokens
        
    Follow [this link](https://docs.microsoft.com/en-us/graph/auth-v2-user#authentication-and-authorization-steps) and obtain the client ID, client secret and refresh token.

 
## Quickstart

### For Choreo user 

Input from Shani

#### Step 1: Select the connector

At Choreo UI, select the `API Calls` tab. Search for `MS Outlook Mail` connector and select it. 

#### Step 2: Configure the connection 

1. Give a name to the connection you are configuring in
`Endpoint Name` text box. 
2. Then you can configure the connection detail. Different connectors will contain different type 
of connections and some may support multiple types of connections as well. Below are some common connection types. 

BearerTokenConfig : In `Token` textbox, configure the bearer token obtained for the service 
Oauth2RefreshTokenGrantConfig: In below text boxes, configure the tokens related to Oauth2 authentication obtained at 
prerequisites section. 

RefreshUrl:  
RefreshToken:
ClientId:
ClientSecret: 


#### Step 3: Invoke connector operation 

1. Click on `Continue To Invoke API` button. It will allow you to select the operation you need to invoke using the connector. 
2. Search and select the operation. Note that the API doc is displayed in a panel at the right hand side. 
3. Fill in the inputs to the operation referring to the API documentation. 
4. Optionally, assign the response of the connector operation to a named variable which you can fill in at `Response Variable Name`. 
5. Click `Save`. You will notice diagram gets updated and you have a variable assigned with connector operation invocation result. 
6. Click `Run & Test` to invoke the connector operation. 


### For Ballerina User
You can use Outlook mail connector in your Ballerina application with the following steps. 

#### Import connector
First, import the `ballerinax/microsoft.outlook.mail` module into the Ballerina project.
```ballerina
import ballerinax/microsoft.outlook.mail;
```

#### Create a new connector instance 
Create a `outlookMail:Configuration` with the OAuth2 tokens obtained, and initialize the connector with it. 
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

#### Invoke connector operation 
1. Now you can use the operations available within the connector. Note that they are in the form of remote operations. 
Following is an example on how to send an email using the connector.

Send an email  

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
    http:Response response = check outlookClient->sendMessage(messageContent);
    log:printInfo(message.toString());
}
```
2. Use `bal run` command to compile and run the Ballerina program. 

## Quick reference 

Input from shani 

* Get detail of a message from a mailbox
 ```ballerina
   mail:Message message = check outlookClient->getMessage("<Message ID>");
   ```
* Send an existing draft
 ```ballerina
    _= check outlookClient->sendDraftMessage("<Draft ID>");
```

**[You can find more samples here](https://github.com/ballerina-platform/module-ballerinax-microsoft.outlook.mail/tree/main/outlookmail/samples)**