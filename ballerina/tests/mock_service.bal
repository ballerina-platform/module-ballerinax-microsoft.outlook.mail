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

import ballerina/http;

listener http:Listener mockEp = check new (9090, config = {host: "localhost"});

int mockIdCounter = 0;
map<MicrosoftGraphMessage> mockMessages = {};
map<MicrosoftGraphMailFolder> mockFolders = {};
map<MicrosoftGraphAttachment> mockAttachments = {};

function generateMockId() returns string {
    mockIdCounter += 1;
    return "mock-id-" + mockIdCounter.toString();
}

service / on mockEp {

    resource function post me/messages(@http:Payload MicrosoftGraphMessage payload) returns MicrosoftGraphMessage {
        string id = generateMockId();
        payload.id = id;
        mockMessages[id] = payload;
        return payload;
    }

    resource function get me/messages() returns MicrosoftGraphMessageCollectionResponse {
        return {value: mockMessages.toArray()};
    }

    resource function get me/messages/[string messageId]() returns MicrosoftGraphMessage|http:NotFound {
        MicrosoftGraphMessage? msg = mockMessages[messageId];
        if msg is () {
            return <http:NotFound>{body: {message: "Message not found: " + messageId}};
        }
        return msg;
    }

    resource function patch me/messages/[string messageId](@http:Payload MicrosoftGraphMessage payload) returns MicrosoftGraphMessage|http:NotFound {
        MicrosoftGraphMessage? existing = mockMessages[messageId];
        if existing is () {
            return <http:NotFound>{body: {message: "Message not found: " + messageId}};
        }
        payload.id = messageId;
        mockMessages[messageId] = payload;
        return payload;
    }

    resource function delete me/messages/[string messageId]() returns http:NoContent {
        if mockMessages.hasKey(messageId) {
            _ = mockMessages.remove(messageId);
        }
        return <http:NoContent>{};
    }

    resource function post me/messages/[string messageId]/send() returns http:Accepted {
        return <http:Accepted>{};
    }

    resource function post me/messages/[string messageId]/copy(@http:Payload json payload) returns MicrosoftGraphMessageResponse|http:NotFound {
        MicrosoftGraphMessage? existing = mockMessages[messageId];
        if existing is () {
            return <http:NotFound>{body: {message: "Message not found: " + messageId}};
        }
        string newId = generateMockId();
        MicrosoftGraphMessage copied = existing.clone();
        copied.id = newId;
        mockMessages[newId] = copied;
        return copied;
    }

    resource function post me/messages/[string messageId]/forward(@http:Payload json payload) returns http:Accepted {
        return <http:Accepted>{};
    }

    // ── Send mail ─────────────────────────────────────────────────────────────

    resource function post me/sendMail(@http:Payload json payload) returns http:Accepted {
        return <http:Accepted>{};
    }

    resource function get me/messages/[string messageId]/attachments() returns MicrosoftGraphAttachmentCollectionResponse {
        MicrosoftGraphAttachment[] atts = [];
        string prefix = messageId + ":";
        foreach var [key, att] in mockAttachments.entries() {
            if key.startsWith(prefix) {
                atts.push(att);
            }
        }
        return {value: atts};
    }

    resource function post me/messages/[string messageId]/attachments(@http:Payload MicrosoftGraphAttachment payload) returns MicrosoftGraphAttachment {
        string id = generateMockId();
        payload.id = id;
        string key = messageId + ":" + id;
        mockAttachments[key] = payload;
        return payload;
    }

    resource function delete me/messages/[string messageId]/attachments/[string attachmentId]() returns http:NoContent {
        string key = messageId + ":" + attachmentId;
        if mockAttachments.hasKey(key) {
            _ = mockAttachments.remove(key);
        }
        return {};
    }

    resource function post me/messages/[string messageId]/attachments/createUploadSession(
            @http:Payload json payload) returns MicrosoftGraphUploadSessionResponse {
        MicrosoftGraphUploadSession session = {
            atOdataType: "#microsoft.graph.uploadSession",
            uploadUrl: "https://outlook.office.com/api/v2.0/users/mock/messages/" + messageId + "/attachments/upload",
            expirationDateTime: "2026-05-27T00:00:00Z",
            nextExpectedRanges: ["0"]
        };
        return session;
    }

    resource function get me/mailFolders() returns MicrosoftGraphMailFolderCollectionResponse {
        return {value: mockFolders.toArray()};
    }

    resource function post me/mailFolders(@http:Payload MicrosoftGraphMailFolder payload) returns MicrosoftGraphMailFolder {
        string id = generateMockId();
        payload.id = id;
        mockFolders[id] = payload;
        return payload;
    }

    resource function get me/mailFolders/[string mailFolderId]() returns MicrosoftGraphMailFolder|http:NotFound {
        MicrosoftGraphMailFolder? folder = mockFolders[mailFolderId];
        if folder is () {
            return <http:NotFound>{body: {message: "Folder not found: " + mailFolderId}};
        }
        return folder;
    }

    resource function delete me/mailFolders/[string mailFolderId]() returns http:NoContent {
        if mockFolders.hasKey(mailFolderId) {
            _ = mockFolders.remove(mailFolderId);
        }
        return <http:NoContent>{};
    }

    resource function get me/mailFolders/[string mailFolderId]/childFolders() returns MicrosoftGraphMailFolderCollectionResponse {
        return {value: []};
    }

    resource function post me/mailFolders/[string mailFolderId]/childFolders(
            @http:Payload MicrosoftGraphMailFolder payload) returns MicrosoftGraphMailFolder {
        string id = generateMockId();
        MicrosoftGraphMailFolder folder = {...payload};
        folder.id = id;
        return folder;
    }
}
