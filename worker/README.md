# worker

layer: 2 (processing). runs pptx timeline transformation and saves outputs.

## responsibilities
- apply timeline transformations
- save processed pptx
- emit progress + errors

## interface
| direction | data        | protocol | peer         |
| --------- | ----------- | -------- | ------------ |
| input     | work item   | queue    | orchestrator |
| output    | pptx output | s3/http  | storage      |

## wbs
```mermaid
graph TB
  subgraph "worker"
    LOAD[Load PPTX]
    XFORM[Apply Timeline]
    SAVE[Save Output]
    EMIT[Emit Progress]
  end
  ORCH[Orchestrator]
  S3[(S3/MinIO)]
  ORCH --> LOAD
  LOAD --> XFORM
  XFORM --> SAVE
  SAVE --> S3
  XFORM --> EMIT
```


