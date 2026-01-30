"""in-memory cache with ttl for pdf bytes"""
import time
import logging
from threading import Lock

logger = logging.getLogger(__name__)


class PdfCache:
    """thread-safe in-memory cache with ttl and max entries"""

    def __init__(self, ttl_seconds: int = 300, max_entries: int = 20):
        self.ttl = ttl_seconds
        self.max_entries = max_entries
        self._entries: dict[str, tuple[float, bytes]] = {}
        self._lock = Lock()

    def _purge_expired(self, now: float):
        """remove expired entries"""
        expired = [key for key, (timestamp, _) in self._entries.items() if now - timestamp > self.ttl]
        for key in expired:
            del self._entries[key]

    def get(self, key: str) -> bytes | None:
        """get cached data by key"""
        now = time.time()
        with self._lock:
            self._purge_expired(now)
            entry = self._entries.get(key)
            return entry[1] if entry else None

    def set(self, key: str, data: bytes):
        """cache data with timestamp"""
        now = time.time()
        with self._lock:
            self._purge_expired(now)
            if len(self._entries) >= self.max_entries and key not in self._entries:
                first_key = next(iter(self._entries))
                del self._entries[first_key]
            self._entries[key] = (now, data)
