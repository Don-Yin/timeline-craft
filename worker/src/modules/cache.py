import time
import logging
from threading import Lock
from typing import Optional, Dict, Tuple

logger = logging.getLogger(__name__)


class PdfCache:
    def __init__(self, ttl_seconds: int = 300, max_entries: int = 20):
        self.ttl = ttl_seconds
        self.max_entries = max_entries
        self._entries: Dict[str, Tuple[float, bytes]] = {}
        self._lock = Lock()

    def _purge_locked(self, now: float):
        # 1. Purge expired items
        expired = [key for key, (timestamp, _) in self._entries.items() if now - timestamp > self.ttl]
        for key in expired:
            del self._entries[key]

    def get(self, key: str) -> Optional[bytes]:
        now = time.time()
        with self._lock:
            self._purge_locked(now)
            entry = self._entries.get(key)
            if entry:
                return entry[1]
            return None

    def set(self, key: str, data: bytes):
        now = time.time()
        with self._lock:
            self._purge_locked(now)
            if len(self._entries) >= self.max_entries and key not in self._entries:
                first_key = next(iter(self._entries))
                del self._entries[first_key]
            self._entries[key] = (now, data)
