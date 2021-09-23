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

class MessageStream {
    private Message[] messageEntries = [];
    int index = 0;
    string nextLink = EMPTY_STRING;
    ConnectionConfig config;
    string? queryParams;

    public isolated function init(json payload, ConnectionConfig config, string? queryParams = ()) returns @tainted error? {
        json[] messages = let var value = payload.value
            in value is json ? <json[]>value : [];
        foreach json message in messages {
            Message msg = check message.cloneWithType(Message);
            self.messageEntries.push(msg);
        }
        self.nextLink = getNextLink(payload);
        self.config = config;
         self.queryParams = queryParams;
    }

    public isolated function next() returns @tainted record {|Message value;|}|error? {
        if (self.index < self.messageEntries.length()) {
            record {|Message value;|} singleRecord = {value: self.messageEntries[self.index]};
            self.index += 1;
            return singleRecord;
        } else if (self.nextLink != EMPTY_STRING && !self.queryParams.toString().includes("$top")) {
            self.index = 0;
            self.messageEntries = check self.fetchMessages(self.nextLink);
            if (self.messageEntries.length() == 0) {
                return;
            }
            record {|Message value;|} singleRecord = {value: self.messageEntries[self.index]};
            self.index += 1;
            return singleRecord;
        }
    }

    isolated function fetchMessages(string nextLink) returns @tainted Message[]|error {
        record {|Message[] messages; string nextLink;|} result = check sendNextRequest(nextLink, self.config);
        self.nextLink = result.nextLink;
        return result.messages;
    }
}

isolated function sendNextRequest(string nextLink, ConnectionConfig config) returns @tainted record {|Message[] messages;
                                  string nextLink;|}|error {
    http:Client tempOutlookClient = check getOutlookClient(config, nextLink);
    json payload = check tempOutlookClient->get(EMPTY_STRING, targetType = json);
    record {|Message[] messages; string nextLink;|} result = {
        messages: check getMessages(payload),
        nextLink: 
        getNextLink(payload)
    };
    return result;
}

isolated function getMessages(json payload) returns @tainted Message[]|error {
    json[] messages = let var value = payload.value
        in value is json ? <json[]>value : [];
    Message[] messageList = [];
    foreach json message in messages {
        Message msg = check message.cloneWithType(Message);
        messageList.push(msg);
    }
    return messageList;
}

isolated function getNextLink(json payload) returns @tainted string {
    map<json> payloadMap = <map<json>>payload;
    return let var value = trap payloadMap.get("@odata.nextLink")
        in value is string ? value : EMPTY_STRING;
}

class AttachmentStream {
    private FileAttachment[] fileAttachmentEntries = [];
    int index = 0;

    public isolated function init(FileAttachment[] attachments) returns @tainted error? {
        self.fileAttachmentEntries = attachments;
    }

    public isolated function next() returns @tainted record {|FileAttachment value;|}|error? {
        if (self.index < self.fileAttachmentEntries.length()) {
            record {|FileAttachment value;|} singleRecord = {value: self.fileAttachmentEntries[self.index]};
            self.index += 1;
            return singleRecord;
        }
    }
}

class MailFolderStream {
    private MailFolder[] mailFolderEntries = [];
    int index = 0;

    public isolated function init(MailFolder[] attachments) returns @tainted error? {
        self.mailFolderEntries = attachments;
    }

    public isolated function next() returns @tainted record {|MailFolder value;|}|error? {
        if (self.index < self.mailFolderEntries.length()) {
            record {|MailFolder value;|} singleRecord = {value: self.mailFolderEntries[self.index]};
            self.index += 1;
            return singleRecord;
        }
    }
}
