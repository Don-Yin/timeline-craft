# api documentation

this directory contains documentation for different api protocols used in the system.

## protocol-specific documentation
| protocol  | detailed guide                           |
| --------- | ---------------------------------------- |
| REST/HTTP | [restful.md](./protocols/rest.md)        |
| gRPC      | [grpc.md](./protocols/grpc.md)           |
| WebSocket | [websocket.md](./protocols/websocket.md) |

## source of truth
| protocol  | language | source of truth                 | derived artifacts                  | workflow       |
| --------- | -------- | ------------------------------- | ---------------------------------- | -------------- |
| REST/HTTP | python   | python code (fastapi)           | -> openapi spec -> swagger ui docs | code first     |
| gRPC      | any      | .proto files (protocol buffers) | -> generated code (python/ts/etc.) | contract first |
| WebSocket | python   | python code (pydantic models)   | models + docstrings = docs         | code first     |

what this means:
- REST: write python code with type hints, docs auto-generate. if docs are wrong, fix the code
- gRPC: write .proto file, code auto-generates for all languages. if code is wrong, fix the .proto
- WebSocket: write python code with pydantic models, use docstrings for documentation

## where to find things

### ports and services
- **docker-compose.yml** = source of truth for all port mappings
- each service defines its port: `ports: ["8000:8000"]`

### documentation
- **REST**: 
  - individual service: `http://localhost:XXXX/docs` (auto-generated)
  - all services: `http://localhost:8080` (swagger ui aggregator)
- **gRPC**: `.proto` files in `docs/proto/` (source of truth, version control these)
- **WebSocket**: pydantic models + docstrings in your code (no separate docs)

### swagger ui aggregator
add to `docker-compose.yml`:
```yaml
swagger-ui:
  image: swaggerapi/swagger-ui
  ports:
    - "8080:8080"
  environment:
    URLS: |
      [
        {name: "API Gateway", url: "http://api-gateway:8000/api/openapi.json"},
        {name: "Auth Service", url: "http://auth-service:8001/api/openapi.json"}
      ]
```
visit `http://localhost:8080` to see all REST apis in one place