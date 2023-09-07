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
import ballerina/http;
import ballerina/test;
import ballerina/io;

configurable string refreshUrl = os:getEnv("REFRESH_URL");
configurable string refreshToken = os:getEnv("REFRESH_TOKEN");
configurable string clientId = os:getEnv("CLIENT_ID");
configurable string clientSecret = os:getEnv("CLIENT_SECRET");

ConnectionConfig configuration = {
    auth: {
        refreshUrl: refreshUrl,
        refreshToken: refreshToken,
        clientId: clientId,
        clientSecret: clientSecret
    }
};

Client outlookClient = check new (configuration);
string createdDraftId = EMPTY_STRING;
string sentMessageId = EMPTY_STRING;
string mailFolderId = EMPTY_STRING;
string searchMailFolderId = EMPTY_STRING;
string attachmentId = EMPTY_STRING;

@test:Config {
    enable: true,
    dependsOn: [tesCreateDraft]
}
function testListMessages() returns error? {
    log:printInfo("outlookClient->listMessages()");
    var output = outlookClient->listMessages(folderId = "drafts", optionalUriParameters = "?$select:\"sender,subject\"&top:2");
    if (output is stream<Message, error?>) {
        int index = 0;
        _ = check output.forEach(function(Message queryResult) {
            index += 1;
        });
        log:printInfo("Total count of records : " + index.toString());
    } else {
        test:assertFail(msg = output.toString());
    }
}

@test:Config {
    enable: true
}
function tesCreateDraft() {
    log:printInfo("outlookClient->tesCreateDraft()");
    DraftMessage draft = {
        subject: "Test Subject",
        importance: "Low",
        body: {
            "contentType": "HTML",
            "content": "Test content <b>ballerina</b>!"
        },
        toRecipients: [
            {
            emailAddress: {
                address: "AdeleV@contoso.onmicrosoft.com",
                name: "Dhanushka"
            }
        }
        ]
    };
    var output = outlookClient->createMessage(draft);
    if (output is Message) {
        log:printInfo(output?.id.toString());
        createdDraftId = <@untainted>output?.id.toString();
    } else {
        test:assertFail(msg = output.toString());
    }
}

@test:Config {
    enable: true,
    dependsOn: [testListMessages]
}
function testGetMessage() {
    log:printInfo("outlookClient->testGetMessage()");
    var output = outlookClient->getMessage(createdDraftId, bodyContentType = "text");
    if (output is Message) {
        log:printInfo(output?.body.toString());
    } else {
        test:assertFail(msg = output.toString());
    }
}

@test:Config {
    enable: true,
    dependsOn: [testGetMessage]
}
function testUpdateMessage() {
    log:printInfo("outlookClient->updateMessage()");
    MessageUpdateContent message = {
        subject: "Test Subject",
        importance: "Low",
        body: {
            "contentType": "HTML",
            "content": "This is sent by Update operation <b>awesome</b>!"
        },
        toRecipients: [
            {
            emailAddress: {
                address: "sandaruwandanushka@gmail.com",
                name: "Dhanushka"
            }
        }
        ]
    };
    var output = outlookClient->updateMessage(createdDraftId, message, "drafts");
    if (output is Message) {
        log:printInfo(output?.subject.toString());
    } else {
        test:assertFail(msg = output.toString());
    }
}

@test:Config {
    enable: true,
    dependsOn: [testUpdateMessage]
}
function testCopyMessage() {
    log:printInfo("outlookClient->testCopyMessage()");
    var output = outlookClient->copyMessage(createdDraftId, "sentitems", "drafts");
    if (output is error) {
        test:assertFail(msg = output.toString());
    }
}

@test:Config {
    enable: true,
    dependsOn: [testCopyMessage]
}
function testForwardMessage() {
    log:printInfo("outlookClient->testForwardMessage()");
    string comment = "test comment for forwarding";
    var output = outlookClient->forwardMessage(createdDraftId, comment, ["danusricom@gmail.com"], "sentItems");
    if (output is error) {
        test:assertFail(msg = output.toString());
    }
}

@test:Config {
    enable: true,
    dependsOn: [testForwardMessage]
}
function testSendExistingDraftMessage() {
    log:printInfo("outlookClient->testSendDraftMessage()");
    var output = outlookClient->sendDraftMessage(createdDraftId);
    if (output is error) {
        test:assertFail(msg = output.toString());
    }
}

@test:Config {
    enable: true,
    dependsOn: [testDeleteAttachment]
}
function testDeleteMessage() returns error? {
    log:printInfo("outlookClient->testDeleteMessage()");
    var output = outlookClient->listMessages(folderId = "sentitems", 
    optionalUriParameters = "?$select:\"sender,subject,hasAttachments\"");
    if (output is stream<Message, error?>) {
        int index = 0;
        _ = check output.forEach(function(Message queryResult) {
            if (queryResult?.hasAttachments == false) {
                createdDraftId = queryResult?.id.toString();
            }
            index = index + 1;
        });
        log:printInfo("Total count of sent messages : " + index.toString());
    }
    var response = outlookClient->deleteMessage(createdDraftId, "sentitems");
    if (response is error) {
        test:assertFail(msg = response.toString());
    }
}

@test:Config {
    enable: true,
    dependsOn: [testSendExistingDraftMessage]
}
function testSendMessage() {
    log:printInfo("outlookClient->testSendMessage()");
    FileAttachment attachment1 = {
        contentBytes: "SGVsbG8gV29ybGQh",
        contentType: "text/plain",
        name: "sample.txt"
    };
    FileAttachment attachment2 = {
        contentBytes: "SGVsbG8gV29ybGQh",
        contentType: "text/plain",
        name: "sample2.txt"
    };
    MessageContent messageReq = {
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
                        address: "dhanushkas@wso2.com",
                        name: "Dhanushka"
                    }
                }
            ],
            attachments: [attachment1, attachment2]
        },
        saveToSentItems: true
    };

    var output = outlookClient->sendMessage(messageReq);
    if (output is error) {
        test:assertFail(msg = output.toString());
    } 
}

@test:Config {
    enable: true
}
function testSendMessageWithoutAttachment() {
    log:printInfo("outlookClient->testSendMessageWithoutAttachments()");
    MessageContent messageReq = {
        message: {
            subject: "Ballerina Test Email Without an Attachment",
            importance: "Low",
            body: {
                "contentType": "HTML",
                "content": "This is sent by sendMessage operation <b>Test</b>!"
            },
            toRecipients: [
                {
                    emailAddress: {
                        address: "dhanushkas@wso2.com",
                        name: "Dhanushka"
                    }
                }
            ]
        },
        saveToSentItems: true
    };

    var output = outlookClient->sendMessage(messageReq);
    if (output is error) {
        test:assertFail(msg = output.toString());
    } 
}

@test:Config {
    enable: true,
    dependsOn: [testSendMessage]
}
function testAddAttachment() returns error? {
    log:printInfo("outlookClient->testAddFileAttachment()");
    var output = outlookClient->listMessages(folderId = "sentitems", 
    optionalUriParameters = "?$select: \"sender,subject,hasAttachments\"");
    if (output is stream<Message, error?>) {
        int index = 0;
        _ = check output.forEach(function(Message queryResult) {
            sentMessageId = queryResult?.id.toString();
            index = index + 1;
        });
        log:printInfo("Total count of sent messages : " + index.toString());
    }
    FileAttachment attachment = {
        contentBytes: "SGVsbG8gV29ybGQh",
        contentType: "text/plain",
        name: "sample3_separate.txt"
    };
    var result = outlookClient->addFileAttachment(sentMessageId, attachment, "sentitems");
    if result is FileAttachment {
        log:printInfo(result.toString());
        attachmentId = result?.id.toString();
    } else {
        test:assertFail(msg = result.toString());
    }    
}

@test:Config {
    enable: true,
    dependsOn: [testAddAttachment]
}
function testListAttachment() returns error? {
    var result = outlookClient->listAttachment(sentMessageId, "sentitems");
    if (result is stream<FileAttachment, error?>) {
        int index = 0;
        _ = check result.forEach(function(FileAttachment queryResult) {
            index = index + 1;
        });
        log:printInfo("Total count of  Attachments : " + index.toString());
    }
}

@test:Config {
    enable: true
}
function testCreateMailFolder() {
    log:printInfo("outlookClient->testCreateMailFolder()");
    var result = outlookClient->createMailFolder("Test_Folder_01", false);
    if result is error {
        test:assertFail(msg = result.toString());
    } else {
        log:printInfo(result.toString());
        mailFolderId = <@untainted>result?.id.toString();
    }
}

@test:Config {
    enable: true,
    dependsOn: [testCreateMailFolder]
}
function testListMailFolders() returns error? {
    log:printInfo("outlookClient->testListMailFolder()");
    var result = outlookClient->listMailFolders(true);
    if (result is stream<MailFolder, error?>) {
        int index = 0;
        _ = check result.forEach(function(MailFolder queryResult) {
            index = index + 1;
        });
        log:printInfo("Total count of  MailFolders : " + index.toString());
    } else {
        test:assertFail(msg = result.toString());
    }
}

@test:Config {
    enable: true,
    dependsOn: [testListMailFolders]
}
function testGetMailFolder() {
    log:printInfo("outlookClient->testGetMailFolder()");
    var result = outlookClient->getMailFolder(mailFolderId);
    if result is MailFolder {
        log:printInfo(result.toString());
    } else {
        test:assertFail(msg = result.toString());
    }

}

@test:Config {
    enable: true,
    dependsOn: [testGetMailFolder]
}
function testCreateChildMailFolder() {
    log:printInfo("outlookClient->testCreateChildMailFolder()");
    var result = outlookClient->createChildMailFolder(mailFolderId, "Test Child Folder", false);
    if result is error {
        test:assertFail(msg = result.toString());
    } else {
        log:printInfo(result.toString());
    }
}

@test:Config {
    enable: true,
    dependsOn: [testCreateChildMailFolder]
}
function testListChildMailFolders() returns error? {
    log:printInfo("outlookClient->testListChildMailFolders()");
    var result = outlookClient->listChildMailFolders(mailFolderId, true);
    if (result is stream<MailFolder, error?>) {
        int index = 0;
        _ = check result.forEach(function(MailFolder queryResult) {
            index = index + 1;
        });
        log:printInfo("Total count of  Child MailFolders : " + index.toString());
    } else {
        test:assertFail(msg = result.toString());
    }
}

@test:Config {
    enable: true,
    dependsOn: [testListChildMailFolders]
}
function testDeleteMailFolder() {
    log:printInfo("outlookClient->testDeleteMailFolder()");
    var result = outlookClient->deleteMailFolder(mailFolderId);
    if result is error {
        test:assertFail(msg = result.toString());
    }
}

@test:Config {
    enable: true
}
function testCreateMailSearchFolder() {
    log:printInfo("outlookClient->testCreateMailSearchFolder()");
    MailSearchFolder mailSearchFolder = {
        displayName: "TestSearch_04",
        includeNestedFolders: true,
        sourceFolderIds: ["inbox"],
        filterQuery: "contains(subject, 'weekly digest')"
    };
    var output = outlookClient->createMailSearchFolder("inbox", mailSearchFolder);
    if output is error {
        test:assertFail(msg = output.toString());
    } else {
        log:printInfo(output.toString());
        searchMailFolderId = <@untainted>output?.id.toString();
        var result = outlookClient->deleteMailFolder(searchMailFolderId);
        if result is error {
            test:assertFail(msg = result.toString());
        }
    }
}

@test:Config {
    enable: true,
    dependsOn: [testListAttachment]
}
function testAddLargeFileAttachment() returns @tainted error? {
    log:printInfo("outlookClient->TestAddLargeFileAttachments");
    stream<io:Block, io:Error?> blockStream = check 
    io:fileReadBlocksAsStream("tests/sample.pdf", 3000000);
    var output = outlookClient->addLargeFileAttachments(sentMessageId, "myFile.pdf", blockStream, fileSize = 10635049);
    if (output is error) {
        test:assertFail(msg = output.toString());
    }
}

@test:Config {
    enable: true,
    dependsOn: [testAddLargeFileAttachment]
}
function testDeleteAttachment() returns @tainted error? {
    log:printInfo("outlookClient->testDeleteAttachment()");
    var output = outlookClient->deleteAttachment(sentMessageId, attachmentId);
    if (output is error) {
        test:assertFail(msg = output.toString());
    }
}

@test:AfterSuite  {}
function afterFunc() returns error? {
    log:printInfo("Removing sent messages");
    var output = outlookClient->listMessages(folderId = "sentitems", 
    optionalUriParameters = "?$select: \"sender,subject,hasAttachments\"&top=2");
    if (output is stream<Message, error?>) {
        _ = check output.forEach(function(Message queryResult) {
            string messageID = queryResult?.id.toString(); 
            http:Response|error result =  outlookClient->deleteMessage(messageID, "sentitems");
            if (result is error) {
                test:assertFail(msg = result.toString());
            }
        });
    }
}
