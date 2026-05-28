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

final mail:Client outlookClient = check new ({
    auth: {
        clientId,
        clientSecret,
        refreshToken,
        refreshUrl
    }
});

public function main() returns error? {
    // Step 1: List unread messages in the inbox to identify new support requests.
    mail:MicrosoftGraphMessageCollectionResponse inboxResponse = check outlookClient->listMessages(
        dollarFilter = "isRead eq false",
        dollarTop = 10,
        dollarSelect = ["id", "subject", "from", "receivedDateTime", "bodyPreview"]
    );
    mail:MicrosoftGraphMessage[] unreadMessages = inboxResponse.value ?: [];
    io:println("Unread messages in inbox: ", unreadMessages.length());

    if unreadMessages.length() == 0 {
        io:println("No unread messages to process.");
        return;
    }

    // Step 2: Fetch full details of the first unread message for review.
    string? firstMsgIdOpt = unreadMessages[0]?.id;
    if firstMsgIdOpt is () {
        io:println("First unread message has no ID, skipping.");
        return;
    }
    string firstMessageId = firstMsgIdOpt;
    mail:MicrosoftGraphMessage fullMessage = check outlookClient->getMessage(firstMessageId);
    io:println("Reviewing message: ", fullMessage?.subject);
    io:println("From: ", fullMessage?.'from);
    io:println("Preview: ", fullMessage?.bodyPreview);

    // Step 3: Mark the reviewed message as read to track triage progress.
    mail:MicrosoftGraphMessage updatedMessage = check outlookClient->updateMessage(firstMessageId, {
        isRead: true
    });
    io:println("Marked as read: ", updatedMessage?.subject, " (isRead: ", updatedMessage?.isRead, ")");

    // Step 4: Create a "Customer Support" folder to organize processed tickets.
    mail:MicrosoftGraphMailFolder supportFolder = check outlookClient->createMailFolder({
        displayName: "Customer Support"
    });
    io:println("Created folder: ", supportFolder?.displayName, " (ID: ", supportFolder?.id, ")");

    // Step 5: Retrieve the newly created folder details to confirm creation.
    string? folderIdOpt = supportFolder?.id;
    if folderIdOpt is () {
        io:println("Created folder has no ID, skipping folder details retrieval.");
        return;
    }
    string folderId = folderIdOpt;
    mail:MicrosoftGraphMailFolder folderDetails = check outlookClient->getMailFolder(folderId);
    io:println("Folder details — Name: ", folderDetails?.displayName,
        ", Total items: ", folderDetails?.totalItemCount,
        ", Unread: ", folderDetails?.unreadItemCount);

    // Step 6: Delete resolved or spam messages to keep the inbox clean.
    // For this example, remove the second unread message if it exists.
    if unreadMessages.length() > 1 {
        string? spamMsgIdOpt = unreadMessages[1]?.id;
        if spamMsgIdOpt is string {
            check outlookClient->deleteMessage(spamMsgIdOpt);
            io:println("Deleted message: ", unreadMessages[1]?.subject);
        }
    }

    io:println("Inbox triage completed.");
}
