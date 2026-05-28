// Copyright (c) 2021 WSO2 Inc. (http://www.wso2.org) All Rights Reserved.
//
// WSO2 Inc. licenses this file to you under the Apache License,
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

import ballerina/log;
import ballerina/os;
import ballerina/test;

configurable boolean isLiveServer = os:getEnv("IS_LIVE_SERVER") == "true";

configurable string clientId = "client-id";
configurable string clientSecret = "client-secret";
configurable string refreshToken = "refresh-token";
configurable string refreshUrl = "refresh-url";
configurable string recipientEmail = "recipient@sample.com";

isolated string createdDraftId = "";
isolated string sentMessageId = "";
isolated string mailFolderId = "";
isolated string searchMailFolderId = "";
isolated string attachmentId = "";

final Client outlookClient = check initClient();

isolated function initClient() returns Client|error {
    if isLiveServer {
        return check new ({
            auth: {
                refreshUrl,
                refreshToken,
                clientId,
                clientSecret
            }
        });
    }
    return check new ({auth: {token: "test-token"}}, "http://localhost:9090");
}

@test:BeforeSuite
isolated function beforeFunc() returns error? {
    // Delete leftover test folder from any previous failed run
    MicrosoftGraphMailFolderCollectionResponse|error folderResponse = outlookClient->listMailFolders(
        dollarFilter = "displayName eq 'Test_Folder_01'");
    if folderResponse is MicrosoftGraphMailFolderCollectionResponse {
        foreach MicrosoftGraphMailFolder folder in (folderResponse.value ?: []) {
            string fid = folder.id ?: "";
            if fid != "" {
                error? result = outlookClient->deleteMailFolder(fid);
                if result is error {
                    log:printError(string `Failed to delete leftover folder: ${fid}`, 'error = result);
                }
            }
        }
    }
}

@test:Config {groups: ["live_test", "mock_test"]}
isolated function testCreateDraft() returns error? {
    MicrosoftGraphItemBody body = {atOdataType: "#microsoft.graph.itemBody", contentType: "html",
        content: "Test content <b>ballerina</b>!"};
    MicrosoftGraphEmailAddress emailAddr = {atOdataType: "#microsoft.graph.emailAddress", address: recipientEmail,
        name: "Person"};
    MicrosoftGraphRecipient recipient = {atOdataType: "#microsoft.graph.recipient", emailAddress: emailAddr};
    MicrosoftGraphMessage draft = {
        atOdataType: "#microsoft.graph.message",
        subject: "Test Subject",
        importance: "low",
        body: body,
        toRecipients: [recipient]
    };
    MicrosoftGraphMessage output = check outlookClient->createDraftMessage(draft);
    test:assertNotEquals(output?.id, (), "Created draft should have an ID");
    test:assertNotEquals(output?.id ?: "", "", "Created draft ID should not be empty");
    test:assertEquals(output?.subject, "Test Subject", "Draft subject should match what was sent");
    lock {
        createdDraftId = output?.id ?: "";
    }
}

@test:Config {dependsOn: [testCreateDraft], groups: ["live_test", "mock_test"]}
isolated function testListMessages() returns error? {
    MicrosoftGraphMessageCollectionResponse response = check outlookClient->listMessages(
        dollarSelect = ["sender", "subject"], dollarTop = 2);
    MicrosoftGraphMessage[] allMessages = response.value ?: [];
    test:assertTrue(allMessages.length() > 0, "Message list should not be empty");
}

@test:Config {dependsOn: [testListMessages], groups: ["live_test", "mock_test"]}
isolated function testGetMessage() returns error? {
    string draftId;
    lock {
        draftId = createdDraftId;
    }
    MicrosoftGraphMessage output = check outlookClient->getMessage(draftId,
        dollarSelect = ["body", "subject"]);
    test:assertNotEquals(output?.id, (), "Retrieved message should have an ID");
    test:assertNotEquals(output?.body, (), "Retrieved message should have a body");
    test:assertEquals(output?.subject, "Test Subject", "Retrieved subject should match original");
}

@test:Config {dependsOn: [testGetMessage], groups: ["live_test", "mock_test"]}
isolated function testUpdateMessage() returns error? {
    string draftId;
    lock {
        draftId = createdDraftId;
    }
    MicrosoftGraphItemBody body = {atOdataType: "#microsoft.graph.itemBody", contentType: "html",
        content: "This is sent by Update operation <b>awesome</b>!"};
    MicrosoftGraphEmailAddress emailAddr = {atOdataType: "#microsoft.graph.emailAddress", address: recipientEmail,
        name: "Person"};
    MicrosoftGraphRecipient recipient = {atOdataType: "#microsoft.graph.recipient", emailAddress: emailAddr};
    MicrosoftGraphMessage message = {
        atOdataType: "#microsoft.graph.message",
        subject: "Test Subject Updated",
        importance: "low",
        body: body,
        toRecipients: [recipient]
    };
    MicrosoftGraphMessage output = check outlookClient->updateMessage(draftId, message);
    test:assertNotEquals(output?.id, (), "Updated message should have an ID");
    test:assertEquals(output?.subject, "Test Subject Updated", "Subject should reflect the update");
}

@test:Config {dependsOn: [testUpdateMessage], groups: ["live_test", "mock_test"]}
isolated function testCopyMessage() returns error? {
    string draftId;
    lock {
        draftId = createdDraftId;
    }
    MicrosoftGraphMessageResponse copy = check outlookClient->copyMessage(draftId, {destinationId: "sentitems"});
    if copy is MicrosoftGraphMessage {
        test:assertNotEquals(copy?.id, (), "Copied message should have an ID");
    }
}

@test:Config {dependsOn: [testCopyMessage], groups: ["live_test", "mock_test"]}
isolated function testForwardMessage() returns error? {
    string draftId;
    lock {
        draftId = createdDraftId;
    }
    MicrosoftGraphEmailAddress fwdEmailAddr = {atOdataType: "#microsoft.graph.emailAddress", address: recipientEmail};
    MicrosoftGraphRecipient fwdRecipient = {atOdataType: "#microsoft.graph.recipient", emailAddress: fwdEmailAddr};
    check outlookClient->forwardMessage(draftId, {
        comment: "test comment for forwarding",
        toRecipients: [fwdRecipient]
    });
}

@test:Config {dependsOn: [testForwardMessage], groups: ["live_test", "mock_test"]}
isolated function testSendExistingDraftMessage() returns error? {
    string draftId;
    lock {
        draftId = createdDraftId;
    }
    check outlookClient->sendDraftMessage(draftId);
}

@test:Config {dependsOn: [testSendExistingDraftMessage], groups: ["live_test", "mock_test"]}
isolated function testSendMessage() returns error? {
    MicrosoftGraphAttachment attachment1 = {atOdataType: "#microsoft.graph.fileAttachment",
        name: "sample.txt", contentType: "text/plain", "contentBytes": "SGVsbG8gV29ybGQh"};
    MicrosoftGraphAttachment attachment2 = {atOdataType: "#microsoft.graph.fileAttachment",
        name: "sample2.txt", contentType: "text/plain", "contentBytes": "SGVsbG8gV29ybGQh"};
    MicrosoftGraphItemBody body = {atOdataType: "#microsoft.graph.itemBody", contentType: "html",
        content: "This is sent by sendMessage operation <b>Test</b>!"};
    MicrosoftGraphEmailAddress emailAddr = {atOdataType: "#microsoft.graph.emailAddress", address: recipientEmail,
        name: "Person"};
    MicrosoftGraphRecipient recipient = {atOdataType: "#microsoft.graph.recipient", emailAddress: emailAddr};
    MicrosoftGraphMessage message = {
        atOdataType: "#microsoft.graph.message",
        subject: "Ballerina Test Email",
        importance: "low",
        body: body,
        toRecipients: [recipient],
        attachments: [attachment1, attachment2]
    };
    check outlookClient->sendMail({Message: message, SaveToSentItems: true});
}

@test:Config {groups: ["live_test", "mock_test"]}
isolated function testSendMessageWithoutAttachment() returns error? {
    MicrosoftGraphItemBody body = {atOdataType: "#microsoft.graph.itemBody", contentType: "html",
        content: "This is sent by sendMessage operation <b>Test</b>!"};
    MicrosoftGraphEmailAddress emailAddr = {atOdataType: "#microsoft.graph.emailAddress", address: recipientEmail,
        name: "Person"};
    MicrosoftGraphRecipient recipient = {atOdataType: "#microsoft.graph.recipient", emailAddress: emailAddr};
    MicrosoftGraphMessage message = {
        atOdataType: "#microsoft.graph.message",
        subject: "Ballerina Test Email Without an Attachment",
        importance: "low",
        body: body,
        toRecipients: [recipient]
    };
    check outlookClient->sendMail({Message: message, SaveToSentItems: true});
}

@test:Config {dependsOn: [testSendMessage], groups: ["live_test", "mock_test"]}
isolated function testAddAttachment() returns error? {
    MicrosoftGraphItemBody attachBody = {atOdataType: "#microsoft.graph.itemBody", contentType: "html",
        content: "Draft message for attachment test"};
    MicrosoftGraphEmailAddress attachEmailAddr = {atOdataType: "#microsoft.graph.emailAddress", address: recipientEmail,
        name: "Person"};
    MicrosoftGraphRecipient attachRecipient = {atOdataType: "#microsoft.graph.recipient",
        emailAddress: attachEmailAddr};
    MicrosoftGraphMessage attachDraft = {
        atOdataType: "#microsoft.graph.message",
        subject: "Attachment Test Draft",
        body: attachBody,
        toRecipients: [attachRecipient]
    };
    MicrosoftGraphMessage createdMsg = check outlookClient->createDraftMessage(attachDraft);
    test:assertNotEquals(createdMsg?.id, (), "Message created for attachment test should have an ID");
    string msgId = createdMsg.id ?: "";
    lock {
        sentMessageId = msgId;
    }
    MicrosoftGraphAttachment attachment = {
        atOdataType: "#microsoft.graph.fileAttachment",
        name: "sample3_separate.txt",
        contentType: "text/plain",
        "contentBytes": "SGVsbG8gV29ybGQh"
    };
    MicrosoftGraphAttachment result = check outlookClient->addAttachment(msgId, attachment);
    test:assertNotEquals(result?.id, (), "Added attachment should have an ID");
    test:assertEquals(result?.name, "sample3_separate.txt", "Attachment name should match");
    lock {
        attachmentId = result?.id ?: "";
    }
}

@test:Config {dependsOn: [testAddAttachment], groups: ["live_test", "mock_test"]}
isolated function testListAttachment() returns error? {
    string msgId;
    lock {
        msgId = sentMessageId;
    }
    MicrosoftGraphAttachmentCollectionResponse response = check outlookClient->listAttachments(msgId);
    MicrosoftGraphAttachment[] allAttachments = response.value ?: [];
    test:assertTrue(allAttachments.length() > 0, "Should have at least one attachment");
}

@test:Config {groups: ["live_test", "mock_test"]}
isolated function testCreateMailFolder() returns error? {
    MicrosoftGraphMailFolder result = check outlookClient->createMailFolder({
        atOdataType: "#microsoft.graph.mailFolder",
        displayName: "Test_Folder_01",
        isHidden: false
    });
    test:assertNotEquals(result?.id, (), "Created mail folder should have an ID");
    test:assertEquals(result?.displayName, "Test_Folder_01", "Folder display name should match");
    lock {
        mailFolderId = result?.id ?: "";
    }
}

@test:Config {dependsOn: [testCreateMailFolder], groups: ["live_test", "mock_test"]}
isolated function testListMailFolders() returns error? {
    MicrosoftGraphMailFolderCollectionResponse response = check outlookClient->listMailFolders(
        includeHiddenFolders = true);
    MicrosoftGraphMailFolder[] allFolders = response.value ?: [];
    test:assertTrue(allFolders.length() > 0, "Mail folder list should not be empty");
}

@test:Config {dependsOn: [testListMailFolders], groups: ["live_test", "mock_test"]}
isolated function testGetMailFolder() returns error? {
    string folderId;
    lock {
        folderId = mailFolderId;
    }
    MicrosoftGraphMailFolder result = check outlookClient->getMailFolder(folderId);
    test:assertEquals(result?.id, folderId, "Returned folder ID should match the requested ID");
    test:assertEquals(result?.displayName, "Test_Folder_01", "Folder display name should match");
}

@test:Config {dependsOn: [testGetMailFolder], groups: ["live_test", "mock_test"]}
isolated function testCreateChildMailFolder() returns error? {
    string folderId;
    lock {
        folderId = mailFolderId;
    }
    MicrosoftGraphMailFolder result = check outlookClient->createChildFolder(folderId, {
        atOdataType: "#microsoft.graph.mailFolder",
        displayName: "Test Child Folder",
        isHidden: false
    });
    test:assertNotEquals(result?.id, (), "Created child folder should have an ID");
    test:assertEquals(result?.displayName, "Test Child Folder", "Child folder display name should match");
}

@test:Config {dependsOn: [testCreateChildMailFolder], groups: ["live_test", "mock_test"]}
isolated function testListChildMailFolders() returns error? {
    string folderId;
    lock {
        folderId = mailFolderId;
    }
    MicrosoftGraphMailFolderCollectionResponse _ = check outlookClient->listChildFolders(folderId,
        includeHiddenFolders = true);
}

@test:Config {dependsOn: [testListChildMailFolders], groups: ["live_test", "mock_test"]}
isolated function testDeleteMailFolder() returns error? {
    string folderId;
    lock {
        folderId = mailFolderId;
    }
    check outlookClient->deleteMailFolder(folderId);
}

@test:Config {groups: ["live_test", "mock_test"]}
isolated function testCreateMailSearchFolder() returns error? {
    MicrosoftGraphMailFolder output = check outlookClient->createMailFolder({
        atOdataType: "#microsoft.graph.mailSearchFolder",
        displayName: "TestSearch",
        isHidden: false,
        "includeNestedFolders": true,
        "sourceFolderIds": ["inbox"],
        "filterQuery": "contains(subject, 'weekly digest')"
    });
    test:assertNotEquals(output?.id, (), "Created search folder should have an ID");
    test:assertEquals(output?.displayName, "TestSearch", "Search folder display name should match");
    string searchFolderId = output?.id ?: "";
    lock {
        searchMailFolderId = searchFolderId;
    }
    check outlookClient->deleteMailFolder(searchFolderId);
}

@test:Config {dependsOn: [testListAttachment], groups: ["live_test"], enable: isLiveServer}
isolated function testAddLargeFileAttachment() returns error? {
    string msgId;
    lock {
        msgId = sentMessageId;
    }
    MicrosoftGraphUploadSessionResponse sessionResponse = check outlookClient->createUploadSession(msgId, {
        attachmentItem: {
            atOdataType: "#microsoft.graph.attachmentItem",
            attachmentType: "file",
            name: "myFile.pdf",
            size: 10635049
        }
    });
    if sessionResponse is MicrosoftGraphUploadSession {
        test:assertNotEquals(sessionResponse?.uploadUrl, (), "Upload session should have a URL");
    }
}

@test:Config {dependsOn: [testAddLargeFileAttachment], groups: ["live_test"], enable: isLiveServer}
isolated function testDeleteAttachment() returns error? {
    string msgId;
    string attId;
    lock {
        msgId = sentMessageId;
    }
    lock {
        attId = attachmentId;
    }
    check outlookClient->deleteAttachment(msgId, attId);
}

@test:Config {dependsOn: [testDeleteAttachment], groups: ["live_test"], enable: isLiveServer}
isolated function testDeleteMessage() returns error? {
    MicrosoftGraphMessageCollectionResponse response = check outlookClient->listMessages(
        dollarSelect = ["sender", "subject", "hasAttachments"]);
    MicrosoftGraphMessage[] noAttachmentMessages = (response.value ?: []).filter(
        msg => msg?.hasAttachments == false);
    if noAttachmentMessages.length() > 0 {
        check outlookClient->deleteMessage(noAttachmentMessages[0]?.id ?: "");
    }
}

@test:AfterSuite {}
isolated function afterFunc() returns error? {
    MicrosoftGraphMessageCollectionResponse response = check outlookClient->listMessages(
        dollarSelect = ["sender", "subject", "hasAttachments"], dollarTop = 2);
    MicrosoftGraphMessage[] allMessages = response.value ?: [];
    foreach MicrosoftGraphMessage msg in allMessages {
        string messageId = msg?.id ?: "";
        error? result = outlookClient->deleteMessage(messageId);
        if result is error {
            log:printError(string `Failed to delete message: ${messageId}`, 'error = result);
        }
    }

    string sentMsgId;
    lock {
        sentMsgId = sentMessageId;
    }
    if sentMsgId != "" {
        error? msgResult = outlookClient->deleteMessage(sentMsgId);
        if msgResult is error {
            log:printError(string `Failed to delete draft message: ${sentMsgId}`, 'error = msgResult);
        }
    }

    string folderId;
    lock {
        folderId = mailFolderId;
    }
    if folderId != "" {
        error? folderResult = outlookClient->deleteMailFolder(folderId);
        if folderResult is error {
            log:printError(string `Failed to delete mail folder: ${folderId}`, 'error = folderResult);
        }
    }
}
