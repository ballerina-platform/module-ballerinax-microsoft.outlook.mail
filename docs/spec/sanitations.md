# Sanitations for OpenAPI Specification

_Authors_: @Nuvindu \
_Reviewers_: @daneshk @ThisaruGuruge \
_Created_: 2026/05/26 \
_Updated_: 2026/05/26 \
_Edition_: Swan Lake

## Overview

This document describes the sanitations applied to the original [Microsoft Outlook mail API specification](https://learn.microsoft.com/en-us/graph/api/resources/mail-api-overview?view=graph-rest-1.0) to improve usability when generating the Ballerina client code.

## Sanitations

The original Microsoft Graph OpenAPI specification is a large, multi-resource document. The specification was reduced and restructured to cover the main Outlook mail operations. The resulting specification was then **flattened and aligned using the `bal openapi` tool** to ensure compatibility with Ballerina client code generation.

The client is generated using the following command:

```bash
bal openapi -i docs/spec/openapi.yaml --mode client --license docs/license.txt -o ballerina
```

### Improve inline response schemas

The original specification contained several unnamed or auto-generated inline response schemas (typically named `InlineResponse` by OpenAPI generators). These were replaced with descriptive, user-friendly names to improve readability of the generated Ballerina types. The renaming is summarized below:

| Original Name | Renamed To | Description |
|---|---|---|
| `InlineResponse200` | `MicrosoftGraphMessageCollectionResponse` | Response containing a collection of messages |
| `InlineResponse200_1` | `MicrosoftGraphMailFolderCollectionResponse` | Response containing a collection of mail folders |
| `InlineResponse200_2` | `MicrosoftGraphAttachmentCollectionResponse` | Response containing a collection of attachments |
| `InlineResponse200_3` | `MicrosoftGraphUploadSessionResponse` | Response for a large file upload session |
| `InlineResponse200_4` | `MicrosoftGraphMessageResponse` | Response for a single message (success or completed) |
| `InlineResponse` | `CompletedResponse` | Generic nullable object response indicating a completed operation |
| `InlineResponse_1` | `UploadSessionCompletedResponse` | Response indicating that an upload session has completed |
| `InlineResponse_2` | `ODataCountResponse` | Response containing an OData count value |
