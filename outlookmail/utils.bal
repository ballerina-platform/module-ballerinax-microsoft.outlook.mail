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
import ballerina/lang.'int;
import ballerina/io;

isolated function sendRequestGET(http:Client httpClient, string resources) returns @tainted json|error {
    return httpClient->get(resources, targetType = json);
}

isolated function getOutlookClient(Configuration config, string baseUrl) returns http:Client|error {
    http:BearerTokenConfig|http:OAuth2RefreshTokenGrantConfig clientConfig = config.clientConfig;
    http:ClientSecureSocket? socketConfig = config?.secureSocketConfig;
    return check new (baseUrl, {
        auth: clientConfig,
        secureSocket: socketConfig
    });
}

isolated function getRecipientListAsRecord(string comment, string[] addressList) returns ForwardParamsList {
    Recipient[] recipients = [];
    foreach string address in addressList {
        EmailAddress emailAddress = {
            address: address
        };
        Recipient recipient = {
            emailAddress: emailAddress
        };
        recipients.push(recipient);
    }
    return {comment: comment, toRecipients: recipients};
}

isolated function uploadByteArray(byte[] file, UploadSession session) returns @tainted error? {
    int currentPosition = 0;
    int size = 3000000;
    boolean isFinalRequest = false;
    UploadSession uploadSession = session;
    http:Client uploadClient = check new (session?.uploadUrl.toString(), {
        http1Settings: {
            chunking: 
        http:CHUNKING_NEVER
        }
    });
    while !isFinalRequest {
        http:Request request = new;
        int endPosition = currentPosition + size;
        if (endPosition > file.length()) {
            endPosition = file.length();
            isFinalRequest = true;
        }
        byte[] sliced = file.slice(currentPosition, endPosition);
        request.setBinaryPayload(sliced);
        request.addHeader("Content-Length", sliced.length().toString());
        request.addHeader("Content-Type", "application/octet-stream");
        request.addHeader("Content-Range", string `bytes ${currentPosition}-${endPosition - 1}/${file.length().toString()}`);
        if (isFinalRequest) {
            http:Response response = check uploadClient->put("", request);
            if (response.statusCode != http:STATUS_CREATED) {
                fail error((check response.getJsonPayload()).toString());
            }
            break;
        }
        uploadSession = check uploadClient->put("", request, targetType = UploadSession);
        currentPosition = check 'int:fromString(uploadSession.nextExpectedRanges[0]);
    }
}

isolated function uploadByteStream(stream<io:Block, error?> fileStream, int? fileSize, UploadSession session) returns 
                                    @tainted error? {
    int currentPosition = 0;
    int maxSize = 3000000;
    boolean isFinalRequest = false;
    UploadSession uploadSession = session;
    http:Client uploadClient = check new (session?.uploadUrl.toString(), {
        http1Settings: {
            chunking: 
        http:CHUNKING_NEVER
        }
    });
    boolean isOver = false;
    while !isOver {
        record {|byte[] & readonly value;|}? byteBlock = check fileStream.next();
        if (byteBlock is ()) {
            isOver = true;
            return;
        } else if (byteBlock.value.length() > maxSize) {
            fail error("Maxinum io:block byte[] size must be smaller than 3MB (3000000 bytes)");
        } else {
            http:Request request = new;
            int endPosition = currentPosition + byteBlock.value.length();
            request.setBinaryPayload(byteBlock.value);
            request.addHeader("Content-Length", byteBlock.value.length().toString());
            request.addHeader("Content-Type", "application/octet-stream");
            request.addHeader("Content-Range", string `bytes ${currentPosition}-${endPosition - 1}/${fileSize.toString()}`);
            http:Response response = check uploadClient->put("", request);
            currentPosition = endPosition;
        }
    }
}

isolated function addOdataFileType(MessageContent|FileAttachment messageContent) returns FileAttachment[] {
    FileAttachment[] attachments = [];
    if (messageContent is FileAttachment) {
        attachments.push(messageContent);
    } else {
        attachments = messageContent?.message?.attachments ?: [];
    }
    foreach int i in 0 ... attachments.length() {
        FileAttachment attachment = attachments.remove(0);
        FileAttachment attachmentTemp = {
            contentBytes: attachment.contentBytes,
            name: attachment.name,
            contentType: attachment.contentType,
            "@odata.type": "#microsoft.graph.fileAttachment"
        };
        if (attachment?.id is string) {
            attachmentTemp.id = attachment?.id.toString();
        }
        if (attachment?.contentId is string) {
            attachmentTemp.contentId = attachment?.contentId.toString();
        }
        if (attachment?.isInline is boolean) {
            attachmentTemp.isInline = <boolean>attachment?.isInline;
        }
        if (attachment?.lastModifiedDateTime is string) {
            attachmentTemp.lastModifiedDateTime = attachment?.lastModifiedDateTime.toString();
        }
        if (attachment?.size is int) {
            attachmentTemp.size = <int>attachment?.size;
        }
        attachments.push(attachmentTemp);
    }
    return attachments;
}

isolated function addChildFolderIds(string[] childFoldersIds) returns string {
    string requestParams = "";
    foreach string ids in childFoldersIds {
        requestParams += "/childFolders/" + ids;
    }
    return requestParams;
}

isolated function getErrorMessage(http:Response response) returns error {
    fail error(STATUS_CODE + response.statusCode.toString() + COMMA + (check response.getJsonPayload()).toString());
}
