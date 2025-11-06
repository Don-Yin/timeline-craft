# api-gateway

layer: 0 (bff/api). routes frontend requests, handles auth, signs uploads, and exposes job/status endpoints.

## responsibilities
- routing and rate limiting
- auth handoff (jwt) and session checks
- pre-signed upload urls for pptx files
- timeline job submit/status apis

## interface
| direction | data                  | protocol   | peer            |
| --------- | --------------------- | ---------- | --------------- |
| input     | http requests         | http/json  | frontend        |
| output    | auth checks           | http/json  | auth-service    |
| output    | signed upload urls    | http/json  | upload-service  |
| output    | job submit/status     | http/json  | orchestrator    |

## wbs
```mermaid
graph TB
  subgraph "api-gateway"
    API[API Gateway]
    RL[Rate Limiter]
    AUTHZ[Auth Proxy]
    SIGN[Upload Signer]
    JOB[Timeline API]
  end
  FE[Frontend]
  AUTH[Auth Service]
  UP[Upload Service]
  ORCH[Orchestrator]
  FE -->|http/websocket| API
  API --> RL
  API --> AUTHZ
  API --> SIGN
  API --> JOB
  AUTHZ -->|jwt validate| AUTH
  SIGN -->|sign| UP
  JOB -->|enqueue/status| ORCH
```


