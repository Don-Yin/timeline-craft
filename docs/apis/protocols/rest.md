# REST/HTTP APIs

## workflow
```
1. write code with type hints/decorators
     ↓
2. code auto-generates openapi spec
     ↓
3. openapi spec generates docs (swagger ui)
```

**source of truth:** code → documentation is derived

## stack

### fastapi
- auto-generates openapi specs from code
- swagger ui at `/api/docs`
- openapi json at `/api/openapi.json`
- redoc at `/api/redoc`

### swagger ui aggregator
- unified docs for all rest services
- single interface to explore apis

## implementation

```python
from fastapi import FastAPI

app = FastAPI(
    title="API Gateway",
    version="1.0.0",
    openapi_url="/api/openapi.json",
    docs_url="/api/docs",
    redoc_url="/api/redoc"
)

@app.get("/health")
async def health():
    """Health check endpoint"""
    return {"status": "ok"}

@app.post("/orders")
async def create_order(symbol: str, quantity: float, price: float):
    """Place a new trading order"""
    return {
        "order_id": "123",
        "symbol": symbol,
        "quantity": quantity,
        "price": price,
        "status": "pending"
    }
```

## docker-compose setup

```yaml
services:
  # rest api documentation aggregator
  swagger-ui:
    image: swaggerapi/swagger-ui
    ports:
      - "8080:8080"  # unified rest docs
    environment:
      URLS: |
        [
          {name: "API Gateway", url: "http://api-gateway:8000/api/openapi.json"},
          {name: "Auth Service", url: "http://auth-service:8001/api/openapi.json"}
        ]
  
  # rest services
  api-gateway:
    build: ./api-gateway
    ports:
      - "8000:8000"  # rest/http
  
  auth-service:
    build: ./auth-service
    ports:
      - "8001:8001"  # rest/http
```

## finding services

check `docker-compose.yml` for port mappings:
```yaml
api-gateway: http://localhost:8000
auth-service: http://localhost:8001
```

each service auto-documents at `/api/docs`

## testing

### swagger ui (interactive)
- aggregator: `http://localhost:8080` - all services
- individual: `http://localhost:XXXX/api/docs` - per service

### command line
```bash
curl http://localhost:8000/health
http POST http://localhost:8000/orders symbol=BTCUSD quantity=1.0 price=50000
http POST http://localhost:8000/orders Authorization:"Bearer token" symbol=BTCUSD
```

## access points
- swagger aggregator: `http://localhost:8080` - all rest api docs
- individual service docs: `http://localhost:XXXX/api/docs`
- openapi specs: `http://localhost:XXXX/api/openapi.json`
