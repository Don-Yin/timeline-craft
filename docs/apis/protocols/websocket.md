# WebSocket APIs

## workflow
```
1. write code with typed schemas
     ↓
2. code auto-generates asyncapi spec
     ↓
3. asyncapi spec generates documentation
```

**source of truth:** code → asyncapi (code-first, like rest)

## stack

### fastapi + pydantic
- write code with type hints
- pydantic models define message schemas
- models + docstrings = documentation

## simple example

```python
from fastapi import FastAPI, WebSocket
from pydantic import BaseModel
from typing import Literal

app = FastAPI()

# define message schemas
class PriceUpdate(BaseModel):
    """server sends this when price changes"""
    event: Literal["price_update"]
    symbol: str
    price: float

class SubscribeRequest(BaseModel):
    """client sends this to subscribe"""
    action: Literal["subscribe"]
    symbol: str

# websocket endpoint
@app.websocket("/ws/trading")
async def trading_websocket(websocket: WebSocket):
    await websocket.accept()
    
    while True:
        # receive from client
        data = await websocket.receive_json()
        request = SubscribeRequest(**data)
        
        # send to client
        update = PriceUpdate(event="price_update", symbol=request.symbol, price=50000.0)
        await websocket.send_json(update.dict())
```

## testing

```bash
# install wscat
npm install -g wscat

# connect
wscat -c ws://localhost:8000/ws/trading

# send message
> {"action": "subscribe", "symbol": "BTCUSD"}

# receive
< {"event": "price_update", "symbol": "BTCUSD", "price": 50000.0}
```
