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
configurable string recipientEmail = ?;
configurable string recipientName = ?;

public function main() returns error? {
    mail:Client outlookClient = check new ({
        auth: {
            clientId: clientId,
            clientSecret: clientSecret,
            refreshToken: refreshToken,
            refreshUrl: refreshUrl
        }
    });

    // Step 1: Create a dedicated mail folder to store weekly report emails.
    mail:MicrosoftGraphMailFolder reportsFolder = check outlookClient->/me/mailFolders.post({
        displayName: "Weekly Status Reports"
    });
    io:println("Created folder: ", reportsFolder?.displayName, " (ID: ", reportsFolder?.id, ")");

    // Step 2: Create a draft message with an HTML-formatted weekly status report.
    mail:MicrosoftGraphItemBody reportBody = {
        contentType: "html",
        content: string `<h2>Weekly Project Status Report</h2>
<p>Dear Team,</p>
<p>Here is this week's project status summary:</p>
<ul>
  <li><strong>Project Alpha:</strong> On track &mdash; 75% complete</li>
  <li><strong>Project Beta:</strong> Needs attention &mdash; deadline approaching</li>
  <li><strong>Project Gamma:</strong> Completed successfully</li>
</ul>
<p>Please review the attached summary for detailed metrics.</p>
<p>Best regards,<br/>Project Management Team</p>`
    };

    mail:MicrosoftGraphRecipient recipient = {
        emailAddress: <mail:MicrosoftGraphEmailAddress>{address: recipientEmail, name: recipientName}
    };

    mail:MicrosoftGraphMessage draft = check outlookClient->/me/messages.post({
        subject: "[Weekly Report] Project Status Update",
        importance: "normal",
        body: reportBody,
        toRecipients: [recipient]
    });

    string draftId = draft?.id ?: "";
    io:println("Draft created (ID: ", draftId, ")");

    // Step 3: Attach a plain-text summary report to the draft.
    // contentBytes is the base64-encoded content of the summary file.
    mail:MicrosoftGraphAttachment attachment = {
        atOdataType: "#microsoft.graph.fileAttachment",
        name: "status-summary.txt",
        contentType: "text/plain",
        isInline: false
    };
    attachment["contentBytes"] = "UHJvamVjdCBTdGF0dXMgU3VtbWFyeQo9PT09PT09PT09PT09PT09CkFscGhhOiAgNzUlIGNvbXBsZXRlCkJldGE6ICAgTmVlZHMgYXR0ZW50aW9uCkdhbW1hOiAgQ29tcGxldGVk";

    mail:MicrosoftGraphAttachment addedAttachment = check outlookClient->/me/messages/[draftId]/attachments.post(attachment);
    io:println("Attached: ", addedAttachment?.name);

    // Step 4: Send the draft to the recipient.
    check outlookClient->/me/messages/[draftId]/send.post();
    io:println("Weekly status report sent successfully!");

    // Step 5: Confirm delivery by listing recent messages from the mailbox.
    mail:MicrosoftGraphMessageCollectionResponse response = check outlookClient->/me/messages.get();
    mail:MicrosoftGraphMessage[] recentMessages = response.value ?: [];
    io:println("Total messages in mailbox: ", recentMessages.length());
}
