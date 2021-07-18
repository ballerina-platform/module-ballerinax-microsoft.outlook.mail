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

# A message in a mailFolder
#
# + 'from - The owner of the mailbox from which the message is sent 
# + ccRecipients - Field Description  
# + bccRecipients - The Bcc: recipients for the message 
# + body - The body of the message. It can be in HTML or text format 
# + bodyPreview - The body of the message. It can be in HTML or text format 
# + categories - The categories associated with the message
# + changeKey - The version of the message  
# + conversationId - The ID of the conversation the email belongs to
# + conversationIndex - Indicates the position of the message within the conversation
# + createdDateTime - The date and time the message was created. The date and time information uses ISO 8601 format and 
# is always in UTC time. For example, midnight UTC on Jan 1, 2014 is 2014-01-01T00:00:00Z.
# + hasAttachments - Indicates whether the message has attachments
# + importance - The importance of the message. The possible values are: low, normal, and high
# + inferenceClassification - he classification of the message for the user, based on inferred relevance or importance, 
# or on an explicit override. The possible values are: focused or other.
# + internetMessageHeader - A collection of message headers defined by RFC5322
# + internetMessageId - The message ID in the format specified by RFC2822  
# + isDeliveryReceiptRequested - Indicates whether a read receipt is requested for the message
# + isDraft - Indicates whether the message is a draft
# + isRead - Indicates whether the message has been read 
# + isReadReceiptRequested - Indicates whether a read receipt is requested for the message
# + lastModifiedDateTime - The date and time the message was last changed. The date and time information uses ISO 8601 
# format and is always in UTC time. For example, midnight UTC on Jan 1, 2014 is 2014-01-01T00:00:00Z
# + parentFolderId - The unique identifier for the message's parent mailFolder
# + receivedDateTime - The date and time the message was received
# + replyTo - The email addresses to use when replying
# + sender - The account that is actually used to generate the message
# + sentDateTime - The date and time the message was sent. The date and time information uses ISO 8601 
# format and is always in UTC time. For example, midnight UTC on Jan 1, 2014 is 2014-01-01T00:00:00Z 
# + subject - The subject of the message  
# + toRecipients - The To: recipients for the message  
# + uniqueBody - The part of the body of the message that is unique to the current message. uniqueBody is not returned 
# by default but can be retrieved for a given message by use of the ?$select=uniqueBody query. It can be in HTML or text 
# format
# + webLink - The URL to open the message in Outlook on the web
# + flag - The flag value that indicates the status, start date, due date, or completion date for the message  
# + id - Unique identifier for the message (note that this value may change if a message is moved or altered) 
public type Message record {
    Recipient 'from?;
    Recipient[] ccRecipients?;
    Recipient[] bccRecipients?;
    Recipient[] replyTo?;
    string subject?;
    ItemBody body?;
    string bodyPreview?;
    string[] categories?;
    string changeKey?;
    string conversationId?;
    string conversationIndex?;
    string createdDateTime?;
    boolean hasAttachments?;
    string importance?;
    string inferenceClassification?;
    InternetMessageHeader[] internetMessageHeader?;
    string internetMessageId?;
    boolean? isDeliveryReceiptRequested?;
    boolean isDraft?;
    boolean isRead?;
    boolean? isReadReceiptRequested?;
    string lastModifiedDateTime?;
    string parentFolderId?;
    string receivedDateTime?;
    Recipient sender?;
    string sentDateTime?;
    Recipient[] toRecipients?;
    ItemBody uniqueBody?;
    string webLink?;
    Flag flag?;
    string id?;
};

# The properties of a message
#
# + message - Message Details 
# + saveToSentItems - Indicates whether the message is saved in "sentitem" or not
public type MessageContent record {|
    MessageRequestBody message;
    boolean saveToSentItems?;
|};

# Represents the message detail to be sent
#
# + subject - The subject of the message
# + toRecipients - The To: recipients for the message
# + ccRecipients - The Cc: recipients for the message 
# + bccRecipients - The Bcc: recipients for the message
# + body - The body of the message. It can be in HTML or text format
# + categories - The categories associated with the message  
# + hasAttachments - Indicates whether the message has attachments
# + importance - The importance of the message. The possible values are: low, normal, and high.
# + inferenceClassification - The classification of the message for the user, based on inferred relevance or importance, 
# or on an explicit override. The possible values are: focused or other.  
# + isDeliveryReceiptRequested - Indicates whether a read receipt is requested for the message.
# + isReadReceiptRequested - Indicates whether a read receipt is requested for the message
# + attachments - The fileAttachment and itemAttachment attachments for the message
public type MessageRequestBody record {|
    string subject?;
    Recipient[] toRecipients?;
    Recipient[] ccRecipients?;
    Recipient[] bccRecipients?;
    ItemBody body?;
    string[] categories?;
    boolean hasAttachments?;
    string importance?;
    string inferenceClassification?;
    boolean? isDeliveryReceiptRequested?;
    boolean? isReadReceiptRequested?;
    FileAttachment[] attachments?;
|};

# Represents a message to be updated
#
# + bccRecipients - The Bcc: recipients for the message
# + ccRecipients - The Cc: recipients for the message 
# + 'from - The owner of the mailbox from which the message is sent 
# + body - The body of the message. It can be in HTML or text format. Find out about safe HTML in a message body
# + categories - The categories associated with the message  
# + flag - The flag value that indicates the status, start date, due date, or completion date for the message
# + importance - The importance of the message. The possible values are: low, normal, and high  
# + inferenceClassification - The classification of the message for the user, based on inferred relevance or importance, 
# or on an explicit override. The possible values are: focused or other.  
# + changeKey - The version of the message
# + conversationId - The ID of the conversation the email belongs to
# + conversationIndex - Indicates the position of the message within the conversation 
# + createdDateTime - The date and time the message was created. The date and time information uses ISO 8601 format and 
# is always in UTC time. For example, midnight UTC on Jan 1, 2014 is 2014-01-01T00:00:00Z.
# + hasAttachments - Indicates whether the message has attachments  
# + internetMessageId - The message ID in the format specified by RFC2822
# + isDeliveryReceiptRequested - Indicates whether a read receipt is requested for the message
# + isReadReceiptRequested - Indicates whether a read receipt is requested for the message
# + isRead - Indicates whether the message has been read
# + replyTo - The email addresses to use when replying
# + sender - The account that is actually used to generate the message  
# + subject - The subject of the message  
# + toRecipients - The To: recipients for the message
public type MessageUpdateContent record {|
    Recipient[] bccRecipients?;
    Recipient[] ccRecipients?;
    Recipient 'from?;
    ItemBody body?;
    string subject?;
    Recipient[] toRecipients?;
    string[] categories?;
    string importance?;
    string inferenceClassification?;
    string changeKey?;
    string conversationId?;
    string conversationIndex?;
    string createdDateTime?;
    boolean hasAttachments?;
    string internetMessageId?;
    boolean isDeliveryReceiptRequested?;
    boolean isReadReceiptRequested?;
    boolean isRead?;
    Recipient[] replyTo?;
    Recipient sender?;
    Flag flag?;
|};

# Represents the list of forward Recipients and the comment
#
# + comment - The comment of the forward message
# + toRecipients - The recipient list  
public type ForwardParamsList record {|
    string comment;
    Recipient[] toRecipients;
|};

# Represents detail of  a draft 
#
# + subject - The subject of the message 
# + toRecipients - The To: recipients for the message
# + ccRecipients - The Cc: recipients for the message  
# + bccRecipients - The Bcc: recipients for the message  
# + replyTo - The email addresses to use when replying 
# + importance - The importance of the message. The possible values are: low, normal, and high
# + body - The body of the message. It can be in HTML or text format
# + parentFolderId - The ID of the parent folder of the message is saved in 
# + categories - The categories associated with the message
# + inferenceClassification - The classification of the message for the user, based on inferred relevance or importance, 
# or on an explicit override. The possible values are: focused or other. 
# + isDeliveryReceiptRequested - Indicates whether a read receipt is requested for the message
# + isReadReceiptRequested - Indicates whether a read receipt is requested for the message
public type DraftMessage record {|
    string subject?;
    Recipient[] toRecipients?;
    Recipient[] ccRecipients?;
    Recipient[] bccRecipients?;
    Recipient[] replyTo?;
    string importance?;
    ItemBody body?;
    string parentFolderId?;
    string[] categories?;
    string inferenceClassification?;
    boolean isDeliveryReceiptRequested?;
    boolean isReadReceiptRequested?;
|};

# Allows setting a flag in an item for the user to follow up on later
#
# + completedDateTime - The date and time that the follow-up was finished
# + dueDateTime - The date and time that the follow up is to be finished. Note: To set the due date, 
# you must also specify the startDateTime; otherwise, you will get a 400 Bad Request response.
# + flagStatus - The date and time that the follow up is to be finished. Note: To set the due date, you must also 
# specify the startDateTime; otherwise, you will get a 400 Bad Request response.
# + startDateTime - The date and time that the follow-up is to begin
public type Flag record {|
    string completedDateTime?;
    string dueDateTime?;
    string flagStatus;
    string startDateTime?;
|};

# A key-value pair that represents an Internet message header, as defined by RFC5322
#
# + name - Represents the key in a key-value pair
# + value - The value in a key-value pair 
public type InternetMessageHeader record {|
    string name;
    string value;
|};

# The name and email address of a contact or message recipient
#
# + address - The email address of the person or entity
# + name - The display name of the person or entity
public type EmailAddress record {|
    string address;
    string name?;
|};

# Represents a message recipient
#
# + emailAddress - the email address of the recipient
public type Recipient record {|
    EmailAddress emailAddress;
|};

# Represents properties of the body of an item, such as a message, event or group post
#
# + content - The content of the item
# + contentType - The type of the content. Possible values are text and html
public type ItemBody record {|
    string content;
    string contentType;
|};

# A file (such as a text file or Word document) attached to a message
#
# + name - The name representing the text that is displayed below the icon representing the embedded attachment.
# This does not need to be the actual file name.
# + contentType - The content type of the attachment
# + contentBytes - The base64-encoded contents of the file
# + contentId - The ID of the attachment in the Exchange store
# + isInline - Set to true if this is an inline attachment  
# + lastModifiedDateTime - The date and time when the attachment was last modified
# + size - The size in bytes of the attachment
# + id - The attachment ID  
public type FileAttachment record {
    string name;
    string contentType;
    string contentBytes;
    string? contentId?;
    boolean isInline?;
    string lastModifiedDateTime?;
    int size?;
    string id?;
};

# A mail folder in a user's mailbox, such as Inbox and Drafts. Mail folders can contain messages, other Outlook items, 
# and child mail folders. Outlook creates certain folders for users by default. Instead of using the corresponding 
# folder id value, for convenience, you can use the well-known folder names from the table below when accessing these 
# folders. For example, you can get the Drafts folder using its well-known name with the following query.
#
# + childFolderCount - The number of immediate child mailFolders in the current mailFolder
# + displayName - The mailFolder's display name
# + id - The mailFolder's unique identifier
# + isHidden - Indicates whether the mailFolder is hidden. This property can be set only when creating the folder
# + parentFolderId - The unique identifier for the mailFolder's parent mailFolder  
# + totalItemCount - The number of items in the mailFolder
# + unreadItemCount - The number of items in the mailFolder marked as unread
# + childFolders - The collection of child folders in the mailFolder
# + messages - The collection of messages in the mailFolder
# + sizeInBytes - Size of the content in bytes
public type MailFolder record {
    int childFolderCount?;
    string displayName;
    string id?;
    boolean isHidden?;
    string parentFolderId?;
    int totalItemCount?;
    int unreadItemCount?;
    MailFolder[] childFolders?;
    Message[] messages?;
    int sizeInBytes?;
};

# A mailSearchFolder is a virtual folder in the user's mailbox that contains all the email items matching specified 
# search criteria.
#
# + displayName - Display name of the folder
# + isSupported - Indicates whether a search folder is editable using REST APIs 
# + includeNestedFolders - Indicates how the mailbox folder hierarchy should be traversed in the search. true means that
# a deep search should be done to include child folders in the hierarchy of each folder explicitly specified in 
# sourceFolderIds. false means a shallow search of only each of the folders explicitly specified in sourceFolderIds.
# + sourceFolderIds - The mailbox folders that should be mined 
# + filterQuery - The OData query to filter the messages 
# + id - The ID of the mail search folder
public type MailSearchFolder record {
    string displayName;
    boolean isSupported?;
    boolean includeNestedFolders?;
    string[] sourceFolderIds?;
    string filterQuery?;
    string id?;
};

# Represents elated content to a message in the form of an attachment
#
# + attachmentType - The attachment type
# + contentType - The MIME type
# + isInline - true if the attachment is an inline attachment; otherwise, false
# + name - The attachment's file name  
# + size - The length of the attachment in bytes
# + lastModifiedDateTime - The Timestamp type represents date and time information using ISO 8601 format and is always 
# in UTC time. For example, midnight UTC on Jan 1, 2014 is 2014-01-01T00:00:00Z
public type AttachmentItemContent record {
    string attachmentType;
    string contentType?;
    boolean isInline?;
    string name;
    string lastModifiedDateTime?;
    int size;

};

# Represents an array of file attachments  
public type FileAttachments FileAttachment[];

# Represents a large size file attachment
#
# + AttachmentItem - Detail of the attachment
public type LargeFileAttachment record {
    AttachmentItemContent AttachmentItem;
};

# Represents a detail of upload session
#
# + expirationDateTime - The date/time of session expires
# + nextExpectedRanges - The next range of bytes to be uploaded 
# + uploadUrl - The URL for the upload requests  
public type UploadSession record {
    string expirationDateTime;
    string[] nextExpectedRanges;
    string uploadUrl?;
};
