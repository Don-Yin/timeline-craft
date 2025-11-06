# upload-service

layer: 1 (uploads). provides pre-signed urls and validates incoming pptx files.

## responsibilities
- create pre-signed s3/minio urls
- size/type validation, optional antivirus
- attach uploads to user/project

## interface
| direction | data               | protocol  | peer       |
| --------- | ------------------ | --------- | ---------- |
| input     | sign upload        | http/json | api-gateway |
| output    | put object         | s3/http   | storage    |

## wbs
```mermaid
graph TB
  subgraph "upload-service"
    SIGN[Signer]
    VALID[Validator]
    AV[Antivirus (opt)]
  end
  API[API Gateway]
  S3[(S3/MinIO)]
  API --> SIGN
  SIGN --> VALID
  VALID --> AV
  VALID --> S3
```


