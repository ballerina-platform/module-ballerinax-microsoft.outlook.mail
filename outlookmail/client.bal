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

import ballerina/http;
import ballerina/io;

# Microsoft Outlook mail API provides capability to access Outlook mail operations related to messages, attachments, drafts, 
# mail folder, and child mail folders.
@display {
    label: "Microsoft Outlook Mail Client",
    iconPath: "MSOutlookMailLogo.svg"
}
public client class Client {
    http:Client httpClient;
    Configuration config;

    # Gets invoked to initialize the `connector`.
    # The connector initialization requires setting the API credentials.
    # Create an [Microsoft Outlook Account](https://outlook.live.com/owa/) and obtain tokens by following 
    # [this guide](https://docs.microsoft.com/en-us/graph/auth-v2-user#authentication-and-authorization-steps)
    #
    # + config - Configuration for client connector
    # + return - If success returns null otherwise returns the relevant error  
    public isolated function init(Configuration config) returns @tainted error? {
        self.httpClient = check getOutlookClient(config, BASE_URL);
        self.config = config;
    }

    # Gets the messages in the signed-in user's mailbox (including the Deleted Items and Clutter folders)
    #
    # + folderId - The ID of the specific folder in the user's mailbox or the name of well-known folders (inbox, 
    # sentitems etc.)
    # + optionalUriParameters - The optional query parameter string
    # (https://docs.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0&tabs=http#optional-query-parameters)
    # + return - If success returns a ballerina stream of message records otherwise the relevant error
    @display {label: "List Messages"}
    isolated remote function listMessages(@display {label: "Folder ID"} string? folderId = (), 
                                        @display {label: "Optional Query Parameters"} string? optionalUriParameters = ()) 
                                        returns @tainted error|stream<Message, error?> {
        string requestParams = optionalUriParameters is () ? EMPTY_STRING : optionalUriParameters;
        requestParams = folderId is string ? (MAIL_FOLDER + folderId + SLASH_MESSAGES + requestParams) : (SLASH_MESSAGES + 
        requestParams);
        json response = check self.httpClient->get(requestParams, targetType = json);
        MessageStream objectInstance = check new (response, self.config, optionalUriParameters);
        stream<Message, error?> finalStream = new (objectInstance);
        return finalStream;
    }

    # Creates a draft of a new message
    #
    # + message - The detail of the draft  
    # + folderId - The mail folder where the draft should be saved in
    # + return - If success returns the newly created draft detail as a message record otherwise the relevant error
    @display {label: "Create Message"}
    isolated remote function createMessage(@display {label: "Draft Message"} DraftMessage message, 
                                           @display {label: "Folder ID"} string? folderId = ()) returns @tainted 
                                           Message|error {
        string requestParams = folderId is string ? (MAIL_FOLDER + folderId + SLASH_MESSAGES) : SLASH_MESSAGES;
        http:Request request = new;
        request.setJsonPayload(message.toJson());
        request.setHeader(CONTENT_LENGTH, ZERO_STRING);
        return check self.httpClient->post(requestParams, request, targetType = Message);
    }

    # Retrieves the properties and relationships of a message
    #
    # + messageId - The ID of the message  
    # + folderId - The ID of the folder where the message is in
    # + optionalUriParameters - The OData Query Parameters to help customize the response
    # (https://docs.microsoft.com/en-us/graph/api/message-get?view=graph-rest-1.0&tabs=http#optional-query-parameters)
    # + bodyContentType - The format of the body and uniqueBody of the specified message (values : html (default), text) 
    # + return - If success returns the requested message detail as a message record otherwise the relevant error
    @display {label: "Get Message"}
    isolated remote function getMessage(@display {label: "Message ID"} string messageId, 
                                        @display {label: "Folder ID"} string? folderId = (), 
                                        @display {label: "Optional Query Parameters"} string? optionalUriParameters = 
                                        (), @display {label: "Body Content Format"} string bodyContentType = "html") 
                                        returns @tainted Message|error {
        string optionalPrams = optionalUriParameters is () ? EMPTY_STRING : optionalUriParameters;
        string requestParams = folderId is string ? (MAIL_FOLDER + folderId) : EMPTY_STRING;
        requestParams += MESSAGES + messageId + FORWARD_SLASH + optionalPrams;
        map<string> headers = {"Prefer": "outlook.body-content-type=" + bodyContentType};
        return check self.httpClient->get(requestParams, headers, targetType = Message);
    }

    # Updates the properties of a message 
    #
    # + messageId - The ID of the message  
    # + message - The message properties to be updated 
    # + folderId - The ID of the folder where the message is in
    # + return - If success returns the updated message detail as a message record otherwise the relevant error
    @display {label: "Update Message"}
    isolated remote function updateMessage(@display {label: "Update Message"} string messageId, 
                                           @display {label: "Message Content"} MessageUpdateContent message, 
                                           @display {label: "Folder ID"} string? folderId = ()) 
                                           returns @tainted Message|error {
        string requestParams = folderId is string ? (MAIL_FOLDER + folderId + MESSAGES) : MESSAGES;
        requestParams += messageId;
        http:Request request = new;
        request.setJsonPayload(message.toJson());
        return check self.httpClient->patch(requestParams, request, targetType = Message);
    }

    # Deletes a message in the specified user's mailbox
    #
    # + messageId - The ID of the message 
    # + folderId - The ID of the folder where the message is in
    # + return - If success returns null otherwise the relevant error
    @display {label: "Delete Message"}
    isolated remote function deleteMessage(@display {label: "Messages ID"} string messageId, 
                                           @display {label: "Folder ID"} string? folderId = ()) returns @tainted 
                                           http:Response|error {
        string requestParams = folderId is string ? (MAIL_FOLDER + folderId + MESSAGES) : MESSAGES;
        requestParams += messageId;
        return check self.httpClient->delete(requestParams);
    }

    # Sends an existing draft message
    #
    # + messageId - The ID of the message 
    # + return - If success returns null otherwise the relevant error
    @display {label: "Send Draft Messages"}
    isolated remote function sendDraftMessage(@display {label: "Message ID"} string messageId) returns @tainted 
                                              http:Response|error {
        string requestParams = MESSAGES + messageId + SEND;
        http:Request request = new;
        request.setHeader(CONTENT_LENGTH, ZERO_STRING);
        return check self.httpClient->post(requestParams, request);
    }

    # Copies a message to a folder
    #
    # + messageId - The ID of the message 
    # + destinationFolderId - The ID of the destination folder  
    # + folderId - The ID of the folder where the message is in
    # + return - If success returns null otherwise the relevant error
    @display {label: "Copy Message"}
    isolated remote function copyMessage(@display {label: "Message ID"} string messageId, 
                                        @display {label: "Destination Folder ID"} string destinationFolderId, 
                                        @display {label: "Folder ID"} string? folderId = ()) returns @tainted 
                                        http:Response|error {
        string requestParams = folderId is string ? (MAIL_FOLDER + folderId + MESSAGES) : MESSAGES;
        requestParams += messageId + COPY;
        return check self.httpClient->post(requestParams, {"destinationId": destinationFolderId});
    }

    # Forwards a message
    #
    # + messageId - The ID of the message  
    # + comment - The comment of the forwarding message   
    # + addressList - The receivers' email list 
    # + folderId - The ID of the folder where the message is in
    # + return - If success returns the sent message detail as a message record otherwise the relevant error
    @display {label: "Forward Message"}
    isolated remote function forwardMessage(@display {label: "Message ID"} string messageId, 
                                            @display {label: "Comment"} string comment, 
                                            @display {label: "Email List"} string[] addressList, 
                                            @display {label: "Folder ID"} string? folderId = ()) 
                                            returns @tainted http:Response|error {
        string requestParams = folderId is string ? (MAIL_FOLDER + folderId + MESSAGES) : MESSAGES;
        requestParams += messageId + FORWARD;
        ForwardParamsList parameterList = getRecipientListAsRecord(comment, addressList);
        return check self.httpClient->post(requestParams, parameterList.toJson());
    }

    # Sends a message 
    #
    # + messageContent - The message content and properties
    # + return - If success returns null otherwise the relevant error
    @display {label: "Send Message"}
    isolated remote function sendMessage(@display {label: "Message"} MessageContent messageContent) returns @tainted 
                                         http:Response|error {
        string requestParams = SEND_MAIL;
        http:Request request = new;
        messageContent.message.attachments = addOdataFileType(messageContent);
        request.setJsonPayload(messageContent.toJson());
        return check self.httpClient->post(requestParams, request);
    }

    # Retrieves a list of attachments
    #
    # + messageId - The ID of the message 
    # + folderId - The ID of the folder
    # + childFolderIds - The IDs of the childFolders respectively
    # + return - If success returns a ballerina stream of file attachment records otherwise the relevant error
    @display {label: "List Attachments"}
    isolated remote function listAttachment(@display {label: "Message ID"} string messageId, 
                                            @display {label: "Folder ID"} string? folderId = (), 
                                            @display {label: "Child Folder ID List"} string[]? childFolderIds = ()) 
                                            returns @tainted stream<FileAttachment, error?>|error {
        string requestParams = folderId is string ? (MAIL_FOLDER + folderId) : EMPTY_STRING;
        requestParams += childFolderIds is () ? EMPTY_STRING : (addChildFolderIds(childFolderIds));
        requestParams += MESSAGES + messageId + SLASH_ATTACHMENTS;
        json response = check self.httpClient->get(requestParams, targetType = json);
        json[] attachmentList = let var value = response.value
            in value is json ? <json[]>value : [];
        FileAttachment[] attachments = [];
        foreach json attachment in attachmentList {
            FileAttachment fileAttachment = check attachment.cloneWithType(FileAttachment);
            attachments.push(fileAttachment);
        }
        AttachmentStream objectInstance = check new (attachments);
        stream<FileAttachment, error?> finalStream = new (objectInstance);
        return finalStream;
    }

    # Creates a new mail folder in the root folder of the user's mailbox
    #
    # + displayName - The display name of the mail folder  
    # + isHidden - Indicates whether the folder should be hidden or not
    # + return - If success returns the created mail folder detail otherwise the relevant error
    @display {label: "Create Mail Folder"}
    isolated remote function createMailFolder(@display {label: "Display Name"} string displayName, 
                                             @display {label: "Is Hidden"} boolean? isHidden = ()) returns @tainted 
                                             MailFolder|error {
        string requestParams = SLASH_MAIL_FOLDERS;
        http:Request request = new;
        request.setJsonPayload({displayName: displayName, isHidden: false});
        return check self.httpClient->post(requestParams, request, targetType = MailFolder);
    }

    # Creates a new child mailFolder
    #
    # + parentFolderId - The ID of the parent folder
    # + displayName - The display name of the child mail folder  
    # + isHidden - Indicates whether the child folder should be hidden or not
    # + return - If success returns the created mail folder detail otherwise the relevant error
    @display {label: "Create Child Mail Folder"}
    isolated remote function createChildMailFolder(@display {label: "Parent Folder ID"} string parentFolderId, 
                                                    @display {label: "Display Name"} string displayName, 
                                                    @display {label: "Is Hidden"} boolean? isHidden = ()) 
                                                    returns @tainted MailFolder|error {
        string requestParams = MAIL_FOLDER + parentFolderId + SLASH_CHILD_FOLDERS;
        json payload = {displayName: displayName, isHidden: false};
        return check self.httpClient->post(requestParams, payload, targetType = MailFolder);
    }

    # Retrieves the details of a message folder
    #
    # + mailFolderId - The ID of the mail folder
    # + return - If success returns the requested mail folder details as a mail folder record otherwise the relevant 
    # error 
    @display {label: "Get Mail Folder"}
    isolated remote function getMailFolder(@display {label: "Mail Folder ID"} string mailFolderId) returns @tainted 
                                           MailFolder|error {
        string requestParams = MAIL_FOLDER + mailFolderId;
        return check self.httpClient->get(requestParams, targetType = MailFolder);
    }

    # Deletes the specified mail folder or search Mail folder
    #
    # + mailFolderId - The ID of the folder by its well-known folder name, if one exists 
    # ( Eg: inbox, sentitems etc. https://docs.microsoft.com/en-us/graph/api/resources/mailfolder?view=graph-rest-1.0)
    # + return - If success returns null otherwise the relevant error
    @display {label: "Delete Mail Folder"}
    isolated remote function deleteMailFolder(@display {label: "Mail Folder ID"} string mailFolderId) returns @tainted 
                                              http:Response|error {
        string requestParams = MAIL_FOLDER + mailFolderId;
        return check self.httpClient->delete(requestParams);
    }

    # Adds an attachment that's smaller than 3MB to a message
    #
    # + messageId - The ID of the message  
    # + attachment - The File attachment detail 
    # + folderId - The ID of the folder where the message is saved in
    # + childFolderIds - The IDs of the child folders
    # + return - If success returns the added file attachment details as a record otherwise the relevant error
    @display {label: "Add File Attachment"}
    isolated remote function addFileAttachment(@display {label: "Message ID"} string messageId, 
                                                @display {label: "File Attachment"} FileAttachment attachment, 
                                                @display {label: "Folder ID"} string? folderId = (), 
                                                @display {label: "Child Folder IDs"} string[]? childFolderIds = ()) 
                                                returns FileAttachment|error {
        string requestParams = folderId is string ? (MAIL_FOLDER + folderId) : EMPTY_STRING;
        requestParams += childFolderIds is () ? EMPTY_STRING : (addChildFolderIds(childFolderIds));
        requestParams += MESSAGES + messageId + SLASH_ATTACHMENTS;
        http:Request request = new;
        FileAttachment formattedAttachment = addOdataFileType(attachment)[0];
        request.setJsonPayload(formattedAttachment.toJson());
        return check self.httpClient->post(requestParams, request, targetType = FileAttachment);
    }

    # Gets the mail folder collection directly under the root folder
    #
    # + includeHiddenFolders - Indicates whether the hidden folder should be included in the collection or not
    # + return - If success returns a ballerina stream of mail folder records otherwise the relevant error
    @display {label: "List Mail Folders"}
    isolated remote function listMailFolders(@display {label: "Include Hidden Folders"} boolean? includeHiddenFolders = 
                                            ()) returns @tainted stream<MailFolder, error?>|error {
        string requestParams = SLASH_MAIL_FOLDERS;
        requestParams += includeHiddenFolders is () ? EMPTY_STRING : (INCLUDE_HIDDEN_FOLDERS + 
        includeHiddenFolders.toString());
        json response = check self.httpClient->get(requestParams, targetType = json);
        json[] mailFolderList = let var value = response.value
            in value is json ? <json[]>value : [];
        MailFolder[] mailFolders = [];
        foreach json mailFolder in mailFolderList {
            MailFolder mailFolderRecord = check mailFolder.cloneWithType(MailFolder);
            mailFolders.push(mailFolderRecord);
        }
        MailFolderStream objectInstance = check new (mailFolders);
        stream<MailFolder, error?> finalStream = new (objectInstance);
        return finalStream;
    }

    # Gets the folder collection under the specified folder.
    #
    # + parentFolderId - The ID of the parent folder 
    # + includeHiddenFolders - Indicates whether the hidden folder should be included in the collection or not
    # + return - If success returns a ballerina stream of mail folder records otherwise the relevant error 
    @display {label: "List Child Mail Folders"}
    isolated remote function listChildMailFolders(@display {label: "Parent Folder ID"} string parentFolderId, 
                                                @display {label: "Include Hidden Folder"} boolean? 
                                                includeHiddenFolders = ()) returns @tainted stream<MailFolder, error?>
                                                |error {
        string requestParams = MAIL_FOLDER + parentFolderId + SLASH_CHILD_FOLDERS;
        requestParams += includeHiddenFolders is () ? EMPTY_STRING : (INCLUDE_HIDDEN_FOLDERS + 
        includeHiddenFolders.toString());
        json response = check self.httpClient->get(requestParams, targetType = json);
        json[] mailFolderList = let var value = response.value
            in value is json ? <json[]>value : [];
        MailFolder[] mailFolders = [];
        foreach json mailFolder in mailFolderList {
            MailFolder mailFolderRecord = check mailFolder.cloneWithType(MailFolder);
            mailFolders.push(mailFolderRecord);
        }
        MailFolderStream objectInstance = check new (mailFolders);
        stream<MailFolder, error?> finalStream = new (objectInstance);
        return finalStream;
    }

    # Creates a new mailSearchFolder in the specified user's mailbox
    #
    # + parentFolderId - The ID of the parent folder 
    # + mailSearchFolder - The details of the mail search folder
    # + return - If success returns the newly created mail search folder details as a MailSearchFolder record otherwise 
    # the relevant error 
    @display {label: "Create Mail Search Folder"}
    isolated remote function createMailSearchFolder(@display {label: "Parent Folder ID"} string parentFolderId, 
                                                    @display {label: "Mail Search Folder"} MailSearchFolder 
                                                    mailSearchFolder) returns @tainted MailSearchFolder|error {
        string requestParams = MAIL_FOLDER + parentFolderId + SLASH_CHILD_FOLDERS;
        http:Request request = new;
        json searchRequest = mailSearchFolder.toJson();
        _ = check searchRequest.mergeJson({"@odata.type": "microsoft.graph.mailSearchFolder"});
        request.setJsonPayload(searchRequest);
        return check self.httpClient->post(requestParams, request, targetType = MailSearchFolder);
    }

    # Creates an upload session that allows an the client to iteratively upload a file that's lager than 3 MB and
    # lesser than 150MB
    #
    # + messageId - The ID of the message  
    # + attachmentName - The name of the attachment in the message  
    # + file - The file path or the array of bytes of the folder or file path  
    # + fileSize - Size of the file is mandatory for io:Block stream uploading
    # + return - If success returns null otherwise the relevant error
    @display {label: "Add large File Attachment"}
    isolated remote function addLargeFileAttachments(@display {label: "Message ID"} string messageId, 
                                                @display {label: "Attachment Name"} string attachmentName, 
                                                @display {label: "File Path Or Content"} stream<io:Block, error?>|
                                                string|byte[] file, @display {label: "File Size In Bytes"} int? fileSize  
                                                = ()) returns @tainted error? {
        byte[] content = [];
        string requestParams = MESSAGES + messageId + UPLOAD_SESSION;
        http:Request request = new;
        request.addHeader(CONTENT_TYPE, JSON_TYPE);
        if (file is stream<io:Block, error?>) {
            AttachmentItemContent attachmentItem = {
                attachmentType: FILE,
                name: attachmentName,
                size: fileSize ?: 0
            };
            request.setJsonPayload({AttachmentItem: attachmentItem}.toJson());
            UploadSession session = check self.httpClient->post(requestParams, request, targetType = UploadSession);
            return check uploadByteStream(file, fileSize, session);
        } else {
            content = file is string ? check io:fileReadBytes(file) : file;
        }
        AttachmentItemContent attachmentItem = {
            attachmentType: FILE,
            name: attachmentName,
            size: content.length()
        };
        request.setJsonPayload({AttachmentItem: attachmentItem}.toJson());
        UploadSession session = check self.httpClient->post(requestParams, request, targetType = UploadSession);
        check uploadByteArray(content, session);
    }

    # Deletes an attachment
    #
    # + messageId - Message ID  
    # + attachmentID - Attachment ID   
    # + folderId - Folder ID where the file is saved in 
    # + childFolderIds - Child folder IDs
    # + return - If success returns null otherwise the relevant error
    isolated remote function deleteAttachment(@display {label: "Message ID"} string messageId, 
                                                @display {label: "File Attachment ID"} string attachmentID, 
                                                @display {label: "Folder ID"} string? folderId = (), 
                                                @display {label: "Child Folder IDs"} string[]? childFolderIds = ()) 
                                                returns error? {
        string requestParams = folderId is string ? (MAIL_FOLDER + folderId) : EMPTY_STRING;
        requestParams += childFolderIds is () ? EMPTY_STRING : (addChildFolderIds(childFolderIds));
        requestParams += MESSAGES + messageId + ATTACHMENTS + attachmentID;
        http:Response response = check self.httpClient->delete(path = requestParams);
        if (response.statusCode != http:STATUS_NO_CONTENT) {
            return getErrorMessage(response);
        }
    }
}

# Represents configuration parameters to create Azure Cosmos DB client.
#
# + clientConfig - OAuth client configuration
# + secureSocketConfig - SSH configuration
@display {label: "Connection config"}
public type Configuration record {
    http:BearerTokenConfig|http:OAuth2RefreshTokenGrantConfig clientConfig;
    http:ClientSecureSocket secureSocketConfig?;
};
