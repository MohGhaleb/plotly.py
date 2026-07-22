"""Hourly zero-traffic cell detector for MSSQL data sources.

The script keeps a 7-day rolling cache of traffic counters to reduce the number of
queries executed on each run.  It aggregates measurements by eNodeB and identifies
cells that have stopped carrying traffic in the most recent hour despite being
active in the last seven days.  When such cells are found an HTML email is sent
through Outlook via win32com.
"""
from __future__ import annotations

import json
import os
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import pyodbc  # type: ignore

try:
    import win32com.client  # type: ignore
except ImportError:  # pragma: no cover - allows running on non-Windows hosts
    win32com = None  # type: ignore

WINDOW = timedelta(days=7)
DEFAULT_CACHE = Path(__file__).with_name("zero_traffic_cache.json")


@dataclass
class Config:
    """Runtime configuration for the monitor."""

    conn_str: str
    table: str
    recipients: Sequence[str]
    subject: str = "Zero traffic cell alert"
    sender: Optional[str] = None
    cache_path: Path = DEFAULT_CACHE


Record = Dict[str, object]
CellKey = Tuple[str, str]


class CacheManager:
    def __init__(self, path: Path) -> None:
        self.path = path

    def load(self) -> List[Record]:
        if not self.path.exists():
            return []
        with self.path.open("r", encoding="utf-8") as fh:
            raw = json.load(fh)
        for row in raw:
            row["timestamp"] = datetime.fromisoformat(row["timestamp"])  # type: ignore[assignment]
        return raw

    def save(self, rows: Iterable[Record]) -> None:
        serialisable: List[Record] = []
        for row in rows:
            payload = dict(row)
            payload["timestamp"] = payload["timestamp"].isoformat()  # type: ignore[index]
            serialisable.append(payload)
        self.path.write_text(json.dumps(serialisable, indent=2), encoding="utf-8")

    def merge(self, existing: List[Record], new_rows: Iterable[Record]) -> List[Record]:
        merged: Dict[Tuple[datetime, str, str], Record] = {}
        for row in existing:
            merged[(row["timestamp"], row["enodeb_id"], row["cell_id"])] = row  # type: ignore[index]
        for row in new_rows:
            merged[(row["timestamp"], row["enodeb_id"], row["cell_id"])] = row  # type: ignore[index]
        if not merged:
            return []
        newest = max(key[0] for key in merged)
        cutoff = newest - WINDOW
        return [row for row in merged.values() if row["timestamp"] >= cutoff]  # type: ignore[index]


class MSSQLClient:
    def __init__(self, conn_str: str, table: str) -> None:
        self.conn_str = conn_str
        self.table = table

    def fetch_since(self, since: datetime) -> List[Record]:
        query = f"""
        SELECT
            DATEADD(hour, DATEDIFF(hour, 0, [SampleTime]), 0) AS SampleHour,
            [eNodeB_Name],
            CAST([eNodeB_ID] AS NVARCHAR(64)) AS eNodeB_ID,
            CAST([Cell_ID] AS NVARCHAR(32)) AS Cell_ID,
            [L.RRC.ConnReq.Att] AS RRC,
            [L.Thrp.bits.DL] AS Thr
        FROM {self.table}
        WHERE [SampleTime] >= ?
        """
        with pyodbc.connect(self.conn_str) as conn:
            cursor = conn.cursor()
            cursor.execute(query, since)
            rows: List[Record] = []
            for record in cursor.fetchall():
                rows.append(
                    {
                        "timestamp": record.SampleHour,
                        "enodeb_name": record.eNodeB_Name,
                        "enodeb_id": record.eNodeB_ID,
                        "cell_id": record.Cell_ID,
                        "rrc": record.RRC,
                        "thr": record.Thr,
                    }
                )
        return rows


class ZeroTrafficMonitor:
    def __init__(self, config: Config) -> None:
        self.config = config
        self.cache = CacheManager(config.cache_path)
        self.client = MSSQLClient(config.conn_str, config.table)

    def run(self) -> bool:
        cached = self.cache.load()
        since = (max(row["timestamp"] for row in cached) - timedelta(hours=1)) if cached else datetime.utcnow() - WINDOW
        fresh = self.client.fetch_since(since)
        merged = self.cache.merge(cached, fresh)
        self.cache.save(merged)
        summary = self._summarise(merged)
        if not summary:
            return False
        html = self._build_html(summary)
        self._send_email(html)
        return True

    def _summarise(self, rows: List[Record]) -> List[Dict[str, object]]:
        if not rows:
            return []
        grouped: Dict[CellKey, List[Record]] = defaultdict(list)
        for row in rows:
            grouped[(row["enodeb_id"], row["cell_id"] )].append(row)  # type: ignore[index]
        latest = max(row["timestamp"] for row in rows)  # type: ignore[index]
        earliest = latest - WINDOW
        enodeb_totals: Dict[str, Dict[str, object]] = {}
        for (enodeb_id, cell_id), history in grouped.items():
            history.sort(key=lambda r: r["timestamp"])  # type: ignore[index]
            enodeb_totals.setdefault(enodeb_id, {"cells": set(), "zero_cells": [], "name": history[-1]["enodeb_name"]})  # type: ignore[index]
            enodeb_totals[enodeb_id]["cells"].add(cell_id)  # type: ignore[index]
            current = next((row for row in history if row["timestamp"] == latest), None)  # type: ignore[index]
            had_traffic = any(
                earliest <= row["timestamp"] < latest
                and ((row.get("rrc") or 0) > 0 or (row.get("thr") or 0) > 0)
                for row in history
            )
            zero_now = current is None or (
                ((current.get("rrc") or 0) == 0) and ((current.get("thr") or 0) == 0)
            )
            if had_traffic and zero_now:
                enodeb_totals[enodeb_id]["zero_cells"].append(cell_id)  # type: ignore[index]
        result: List[Dict[str, object]] = []
        for enodeb_id, data in enodeb_totals.items():
            zero_cells = sorted(data["zero_cells"])  # type: ignore[index]
            if not zero_cells:
                continue
            result.append(
                {
                    "enodeb_id": enodeb_id,
                    "enodeb_name": data["name"],
                    "total_cells": len(data["cells"]),  # type: ignore[arg-type]
                    "zero_count": len(zero_cells),
                    "zero_list": ", ".join(zero_cells),
                }
            )
        return sorted(result, key=lambda item: item["zero_count"], reverse=True)

    def _build_html(self, summary: Sequence[Dict[str, object]]) -> str:
        rows = "".join(
            f"<tr><td>{row['enodeb_name']}</td><td>{row['enodeb_id']}</td>"
            f"<td>{row['total_cells']}</td><td>{row['zero_count']}</td><td>{row['zero_list']}</td></tr>"
            for row in summary
        )
        return (
            "<html><body>"
            "<p>Zero traffic cells detected in the last hour:</p>"
            "<table border='1' cellspacing='0' cellpadding='4'>"
            "<thead><tr>"
            "<th>eNodeB Name</th><th>eNodeB ID</th><th>Total Cells</th><th>Zero Traffic Cells</th><th>Zero Cell List</th>"
            "</tr></thead><tbody>"
            f"{rows}" "</tbody></table></body></html>"
        )

    def _send_email(self, html: str) -> None:
        if win32com is None:  # pragma: no cover - non Windows fallback
            raise RuntimeError("win32com is required to send email on Windows")
        outlook = win32com.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        if self.config.sender:
            mail.SentOnBehalfOfName = self.config.sender
        mail.To = ";".join(self.config.recipients)
        mail.Subject = self.config.subject
        mail.HTMLBody = html
        mail.Send()


def build_config() -> Config:
    return Config(
        conn_str=os.environ["MSSQL_CONN_STRING"],
        table=os.environ.get("MSSQL_TABLE", "dbo.CellTrafficHourly"),
        recipients=[addr.strip() for addr in os.environ["ALERT_RECIPIENTS"].split(",")],
        subject=os.environ.get("ALERT_SUBJECT", "Zero traffic cell alert"),
        sender=os.environ.get("ALERT_SENDER"),
    )


def main() -> None:
    monitor = ZeroTrafficMonitor(build_config())
    if monitor.run():
        print("Zero traffic cells detected and email sent.")
    else:
        print("No zero traffic cells in the latest hour.")


if __name__ == "__main__":
    main()
