# webhook-service

layer: 0 (events). ingests external webhooks (stripe, status) and forwards internal events.

## responsibilities
- receive stripe events securely
- idempotent processing, retries, dlq
- notify api/fe of billing or job status changes

## interface
| direction | data          | protocol  | peer          |
| --------- | ------------- | --------- | ------------- |
| input     | stripe events | http/json | stripe        |
| output    | updates       | http/json | api-gateway   |

## wbs
```mermaid
graph TB
  subgraph "webhook-service"
    RCV[Receiver]
    VER[Signature Verify]
    IDEMP[Idempotency Store]
    DISP[Dispatcher]
  end
  STRIPE[Stripe]
  API[API Gateway]
  STRIPE --> RCV
  RCV --> VER
  VER --> IDEMP
  IDEMP --> DISP
  DISP --> API
```


