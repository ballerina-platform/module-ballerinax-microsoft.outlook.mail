import ballerina/log;
import ballerina/os;
import ballerina/test;
import ballerina/io;

configurable string refreshUrl = os:getEnv("REFRESH_URL");
configurable string refreshToken = os:getEnv("REFRESH_TOKEN");
configurable string clientId = os:getEnv("CLIENT_ID");
configurable string clientSecret = os:getEnv("CLIENT_SECRET");

Configuration configuration = {
    clientConfig: {
        refreshUrl: refreshUrl,
        refreshToken: refreshToken,
        clientId: clientId,
        clientSecret: clientSecret
    }
};

Client oneDriveClient = check new (configuration);
string messageId = "";
string createdDraftId = "";
string sentMessageId = "";
string mailFolderId = "";
string searchMailFolderId = "";
string attachmentId = "";

@test:Config {
    enable: true,
    dependsOn: [tesCreateDraft]
}
function testListMessages() {
    log:printInfo("oneDriveClient->listMessages()");
    var output = oneDriveClient->listMessages(folderId = "drafts", optionalUriParameters = "?$select:\"sender,subject\"&top:2");
    if (output is stream<Message, error?>) {
        int index = 0;
        error? e = output.forEach(function(Message queryResult) {
            index = index + 1;
            messageId = queryResult?.id.toString();
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
    log:printInfo("oneDriveClient->tesCreateDraft()");
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
    var output = oneDriveClient->createMessage(draft);
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
    log:printInfo("oneDriveClient->testGetMessage()");
    var output = oneDriveClient->getMessage(messageId, bodyContentType = "text");
    if (output is Message) {
        log:printInfo(output?.body.toString());
    } else {
        test:assertFail(msg = output.toString());
    }
}

@test:Config {
    enable: true,
    dependsOn: [testListMessages]
}
function testUpdateMessage() {
    log:printInfo("oneDriveClient->updateMessage()");
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
    var output = oneDriveClient->updateMessage(messageId, message, "drafts");
    if (output is Message) {
        log:printInfo(output?.subject.toString());
    } else {
        test:assertFail(msg = output.toString());
    }
}

@test:Config {
    enable: true,
    dependsOn: [tesCreateDraft]
}
function testCopyMessage() {
    log:printInfo("oneDriveClient->testCopyMessage()");
    var output = oneDriveClient->copyMessage(createdDraftId, "sentitems", "drafts");
    if (output is error) {
        test:assertFail(msg = output.toString());
    }
}

@test:Config {
    enable: true,
    dependsOn: [testCopyMessage]
}
function testForwardMessage() {
    log:printInfo("oneDriveClient->testForwardMessage()");
    string comment = "test comment for forwarding";
    var output = oneDriveClient->forwardMessage(createdDraftId, comment, ["danusricom@gmail.com"], "sentItems");
    if (output is error) {
        test:assertFail(msg = output.toString());
    }
}

@test:Config {
    enable: true,
    dependsOn: [testCopyMessage]
}
function testSendExistingDraftMessage() {
    log:printInfo("oneDriveClient->testSendDraftMessage()");
    var output = oneDriveClient->sendDraftMessage(createdDraftId);
    if (output is error) {
        test:assertFail(msg = output.toString());
    }
}

@test:Config {
    enable: true,
    dependsOn: [testUpdateMessage, testGetMessage, testCopyMessage, testForwardMessage, testSendExistingDraftMessage]
}
function testDeleteMessage() {
    log:printInfo("oneDriveClient->testDeleteMessage()");
    var output = oneDriveClient->listMessages(folderId = "sentitems", 
    optionalUriParameters = "?$select:\"sender,subject,hasAttachments\"");
    if (output is stream<Message, error>) {
        int index = 0;
        error? e = output.forEach(function(Message queryResult) {
            if (queryResult?.hasAttachments == false) {
                messageId = queryResult?.id.toString();
            }
            index = index + 1;
        });
        log:printInfo("Total count of sent messages : " + index.toString());
    }
    var response = oneDriveClient->deleteMessage(messageId, "sentitems");
    if (response is error) {
        test:assertFail(msg = response.toString());
    }
}

@test:Config {
    enable: true,
    dependsOn: [testListMessages]
}
function testSendMessage() {
    log:printInfo("oneDriveClient->testSendMessage()");
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

    var output = oneDriveClient->sendMessage(messageReq);
    if (output is error) {
        test:assertFail(msg = output.toString());
    } else {

    }
}

@test:Config {
    enable: true
//dependsOn: [testSendMessage]
}
function testListAttachment() {
    log:printInfo("oneDriveClient->testListAttachment()");
    var output = oneDriveClient->listMessages(folderId = "sentitems", 
    optionalUriParameters = "?$select: \"sender,subject,hasAttachments\"");
    if (output is stream<Message, error>) {
        int index = 0;
        error? e = output.forEach(function(Message queryResult) {
            if (queryResult?.hasAttachments == true) {
                sentMessageId = queryResult?.id.toString();
            }
            index = index + 1;
        });
        log:printInfo("Total count of sent messages : " + index.toString());
    }
    var result = oneDriveClient->listAttachment(sentMessageId, "sentitems");
    if (result is stream<FileAttachment, error?>) {
        int index = 0;
        error? e = result.forEach(function(FileAttachment queryResult) {
            index = index + 1;
        });
        log:printInfo("Total count of  Attachments : " + index.toString());
    }
}

@test:Config {
    enable: true,
    dependsOn: [testListAttachment]
}
function testAddFileAttachment() {
    log:printInfo("oneDriveClient->testAddFileAttachment()");
    FileAttachment attachment = {
        contentBytes: "SGVsbG8gV29ybGQh",
        contentType: "text/plain",
        name: "sample3_separate.txt"
    };
    var result = oneDriveClient->addFileAttachment(sentMessageId, attachment, "sentitems");
    if result is FileAttachment {
        log:printInfo(result.toString());
        attachmentId = result?.id.toString();
    } else {
        test:assertFail(msg = result.toString());
    }
}

@test:Config {
    enable: true
}
function testCreateMailFolder() {
    log:printInfo("oneDriveClient->testCreateMailFolder()");
    var result = oneDriveClient->createMailFolder("Test_Folder_01", false);
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
function testListMailFolders() {
    log:printInfo("oneDriveClient->testListMailFolder()");
    var result = oneDriveClient->listMailFolders(true);
    if (result is stream<MailFolder, error?>) {
        int index = 0;
        error? e = result.forEach(function(MailFolder queryResult) {
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
    log:printInfo("oneDriveClient->testGetMailFolder()");
    var result = oneDriveClient->getMailFolder(mailFolderId);
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
    log:printInfo("oneDriveClient->testCreateChildMailFolder()");
    var result = oneDriveClient->createChildMailFolder(mailFolderId, "Test Child Folder", false);
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
function testListChildMailFolders() {
    log:printInfo("oneDriveClient->testListChildMailFolders()");
    var result = oneDriveClient->listChildMailFolders(mailFolderId, true);
    if (result is stream<MailFolder, error?>) {
        int index = 0;
        error? e = result.forEach(function(MailFolder queryResult) {
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
    log:printInfo("oneDriveClient->testDeleteMailFolder()");
    var result = oneDriveClient->deleteMailFolder(mailFolderId);
    if result is error {
        test:assertFail(msg = result.toString());
    }
}

@test:Config {
    enable: true
}
function testCreateMailSearchFolder() {
    log:printInfo("oneDriveClient->testCreateMailSearchFolder()");
    MailSearchFolder mailSearchFolder = {
        displayName: "TestSearch_03",
        includeNestedFolders: true,
        sourceFolderIds: ["inbox"],
        filterQuery: "contains(subject, 'weekly digest')"
    };
    var output = oneDriveClient->createMailSearchFolder("inbox", mailSearchFolder);
    if output is error {
        test:assertFail(msg = output.toString());
    } else {
        log:printInfo(output.toString());
        searchMailFolderId = <@untainted>output?.id.toString();
        var result = oneDriveClient->deleteMailFolder(searchMailFolderId);
        if result is error {
            test:assertFail(msg = result.toString());
        }
    }
}

@test:Config {
    enable: true,
    dependsOn: [tesCreateDraft]
}
function testAddLargeFileAttachment() returns @tainted error? {
    log:printInfo("oneDriveClient->TestAddLargeFileAttachments");
    stream<io:Block, io:Error?> blockStream = check 
    io:fileReadBlocksAsStream("outlookmail/tests/sample.pdf", 3000000);
    var output = oneDriveClient->addLargeFileAttachments(createdDraftId, "myFile.pdf", blockStream, fileSize = 10635049);
    if (output is error) {
        test:assertFail(msg = output.toString());
    }
}

@test:Config {
    enable: true,
    dependsOn: [testAddFileAttachment]
}
function testDeleteAttachment() returns @tainted error? {
    log:printInfo("oneDriveClient->testDeleteAttachment()");
    var output = oneDriveClient->deleteAttachment(sentMessageId, attachmentId);
    if (output is error) {
        test:assertFail(msg = output.toString());
    }
}
