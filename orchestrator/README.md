# orchestrator

layer: 1 (orchestration). schedules and tracks pptx processing jobs.

## responsibilities
- enqueue timeline jobs
- track job state and retries
- route work to worker service

## interface
| direction | data         | protocol  | peer        |
| --------- | ------------ | --------- | ----------- |
| input     | job requests | http/json | api-gateway |
| output    | work items   | queue     | worker      |

## wbs
```mermaid
graph TB
  subgraph "orchestrator"
    APIH[API Handler]
    QUEUE[Queue]
    SCHED[Scheduler]
    ROUTE[Worker Router]
  end
  API[API Gateway]
  WRK[Worker]
  API --> APIH
  APIH --> QUEUE
  SCHED --> QUEUE
  QUEUE --> ROUTE
  ROUTE --> WRK
```


