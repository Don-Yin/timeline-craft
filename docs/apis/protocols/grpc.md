# gRPC APIs

## workflow
```
1. write .proto file (contract/interface) ← source of truth
     ↓
2. proto generates code (python/go/etc.)
     ↓
3. implement the generated interface
```

**source of truth:** .proto file → code is derived
**languages:** any (python, typescript, go, java, etc.) - proto is language-agnostic

## define proto file
```protobuf
// trading_engine.proto
syntax = "proto3";

package trading;

service TradingEngine {
  rpc PlaceOrder(OrderRequest) returns (OrderResponse);
  rpc GetPortfolio(PortfolioRequest) returns (Portfolio);
  rpc StreamPositions(StreamRequest) returns (stream Position);
}

message OrderRequest {
  string symbol = 1;
  double quantity = 2;
  double price = 3;
  OrderSide side = 4;
}

message OrderResponse {
  string order_id = 1;
  OrderStatus status = 2;
  string message = 3;
}

enum OrderSide {
  BUY = 0;
  SELL = 1;
}

enum OrderStatus {
  PENDING = 0;
  FILLED = 1;
  REJECTED = 2;
}
```

## generate code
```bash
# python
protoc --python_out=. --grpc_python_out=. trading_engine.proto

# other languages supported (go, java, c++, etc.)
# see: https://grpc.io/docs/languages/
```

## implement service
```python
# trading_engine.py
import grpc
from concurrent import futures
import trading_engine_pb2_grpc
import trading_engine_pb2

class TradingEngineServicer(trading_engine_pb2_grpc.TradingEngineServicer):
    def PlaceOrder(self, request, context):
        # implement business logic
        return trading_engine_pb2.OrderResponse(
            order_id="order_123",
            status=trading_engine_pb2.PENDING,
            message="Order placed successfully"
        )
    
    def GetPortfolio(self, request, context):
        # implementation
        pass
    
    def StreamPositions(self, request, context):
        # server streaming implementation
        while True:
            position = get_current_position()
            yield position

# start grpc server
server = grpc.server(futures.ThreadPoolExecutor(max_workers=10))
trading_engine_pb2_grpc.add_TradingEngineServicer_to_server(
    TradingEngineServicer(), server
)
server.add_insecure_port('[::]:50051')

# enable grpc reflection for runtime introspection
from grpc_reflection.v1alpha import reflection
reflection.enable_server_reflection(
    trading_engine_pb2.DESCRIPTOR.services_by_name.values(),
    server
)

server.start()
server.wait_for_termination()
```

## docker-compose setup

```yaml
services:
  # grpc services (internal, high-performance)
  trading-engine:
    build: ./trading-engine
    ports:
      - "50051:50051"  # grpc
    # optionally expose rest admin endpoint
    # - "8007:8007"
  
  strategy-executor:
    build: ./strategy-executor
    ports:
      - "50052:50052"  # grpc
```

## finding services

check `docker-compose.yml` for port mappings:
```yaml
trading-engine: localhost:50051
strategy-executor: localhost:50052
```

## documentation

store `.proto` files in `docs/proto/` and version control them - they are your source of truth

## ci/cd integration

```bash
# compile proto files for all languages
protoc --python_out=. --grpc_python_out=. docs/proto/*.proto

# test grpc service with reflection
grpcurl -plaintext localhost:50051 list
grpcurl -plaintext localhost:50051 describe TradingEngine

# generate html documentation
protoc --doc_out=docs/grpc --doc_opt=html,index.html docs/proto/*.proto

# generate markdown documentation
protoc --doc_out=docs/grpc --doc_opt=markdown,api.md docs/proto/*.proto
```
