import threading
import time
from collections import defaultdict


class RequestTracker:

    def __init__(self):
        self.lock = threading.Lock()
        self.records = []

    def log_request(self, source: str = ""):
        with self.lock:
            self.records.append({
                "worker": threading.current_thread().name,
                "source": source,
                "timestamp": time.time()
            })

    def summary(self) -> dict:
        with self.lock:
            if not self.records:
                return {"total_requests": 0, "per_worker": {}, "per_source": {}}

            per_worker = defaultdict(list)
            per_source_count = defaultdict(int)

            for r in self.records:
                per_worker[r["worker"]].append(r["timestamp"])
                per_source_count[r["source"] or "unknown"] += 1

            worker_stats = {}
            for worker, timestamps in per_worker.items():
                timestamps.sort()
                duration_min = (timestamps[-1] - timestamps[0]) / 60 if len(timestamps) > 1 else 0
                req_count = len(timestamps)
                req_per_min = req_count / duration_min if duration_min > 0 else req_count
                worker_stats[worker] = {
                    "requests": req_count,
                    "duration_min": round(duration_min, 2),
                    "req_per_min": round(req_per_min, 2)
                }

            all_ts = sorted(r["timestamp"] for r in self.records)
            total_duration_min = (all_ts[-1] - all_ts[0]) / 60 if len(all_ts) > 1 else 0
            total_req_per_min = len(all_ts) / total_duration_min if total_duration_min > 0 else len(all_ts)

            return {
                "total_requests": len(self.records),
                "total_duration_min": round(total_duration_min, 2),
                "total_req_per_min": round(total_req_per_min, 2),
                "per_worker": worker_stats,
                "per_source": dict(per_source_count)
            }

    def save(self, filepath: str):
        import json
        stats = self.summary()
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(stats, f, ensure_ascii=False, indent=2)
        return stats


tracker = RequestTracker()