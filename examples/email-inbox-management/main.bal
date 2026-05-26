// Copyright (c) 2026, WSO2 LLC. (http://www.wso2.com).
//
// WSO2 LLC. licenses this file to you under the Apache License,
// Version 2.0 (the "License"); you may not use this file except
// in compliance with the License.
// You may obtain a copy of the License at
//
// http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing,
// software distributed under the License is distributed on an
// "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
// KIND, either express or implied.  See the License for the
// specific language governing permissions and limitations
// under the License.

import ballerina/io;
import ballerinax/microsoft.outlook.mail;

configurable string refreshUrl = ?;
configurable string refreshToken = ?;
configurable string clientId = ?;
configurable string clientSecret = ?;

public function main() returns error? {
    mail:Client outlookClient = check new ({
        auth: {
            clientId: clientId,
            clientSecret: clientSecret,
            refreshToken: refreshToken,
            refreshUrl: refreshUrl
        }
    });

    // Step 1: List unread messages in the inbox to identify new support requests.
    mail:MicrosoftGraphMessageCollectionResponse inboxResponse = check outlookClient->/me/messages.get(
        queries = {
            dollarFilter: "isRead eq false",
            dollarTop: 10,
            dollarSelect: ["id", "subject", "from", "receivedDateTime", "bodyPreview"]
        }
    );
    mail:MicrosoftGraphMessage[] unreadMessages = inboxResponse.value ?: [];
    io:println("Unread messages in inbox: ", unreadMessages.length());

    if unreadMessages.length() == 0 {
        io:println("No unread messages to process.");
        return;
    }

    // Step 2: Fetch full details of the first unread message for review.
    string firstMessageId = unreadMessages[0]?.id ?: "";
    mail:MicrosoftGraphMessage fullMessage = check outlookClient->/me/messages/[firstMessageId].get();
    io:println("Reviewing message: ", fullMessage?.subject);
    io:println("From: ", fullMessage?.'from);
    io:println("Preview: ", fullMessage?.bodyPreview);

    // Step 3: Mark the reviewed message as read to track triage progress.
    mail:MicrosoftGraphMessage updatedMessage = check outlookClient->/me/messages/[firstMessageId].patch({
        isRead: true
    });
    io:println("Marked as read: ", updatedMessage?.subject, " (isRead: ", updatedMessage?.isRead, ")");

    // Step 4: Create a "Customer Support" folder to organize processed tickets.
    mail:MicrosoftGraphMailFolder supportFolder = check outlookClient->/me/mailFolders.post({
        displayName: "Customer Support"
    });
    io:println("Created folder: ", supportFolder?.displayName, " (ID: ", supportFolder?.id, ")");

    // Step 5: Retrieve the newly created folder details to confirm creation.
    string folderId = supportFolder?.id ?: "";
    mail:MicrosoftGraphMailFolder folderDetails = check outlookClient->/me/mailFolders/[folderId].get();
    io:println("Folder details — Name: ", folderDetails?.displayName,
        ", Total items: ", folderDetails?.totalItemCount,
        ", Unread: ", folderDetails?.unreadItemCount);

    // Step 6: Delete resolved or spam messages to keep the inbox clean.
    // For this example, remove the second unread message if it exists.
    if unreadMessages.length() > 1 {
        string spamMessageId = unreadMessages[1]?.id ?: "";
        check outlookClient->/me/messages/[spamMessageId].delete();
        io:println("Deleted message: ", unreadMessages[1]?.subject);
    }

    io:println("Inbox triage completed.");
}
