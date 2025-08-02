import datetime as dt
import time
from typing import Tuple

import numpy as np
import pandas as pd
import plotly.graph_objects as go

try:
    import pyodbc
except Exception:  # pragma: no cover - optional dependency
    pyodbc = None

try:  # pragma: no cover - only available on Windows
    import win32com.client as win32
except Exception:  # pragma: no cover - optional dependency
    win32 = None

DB_CONNECTION_STRING = "DRIVER={ODBC Driver 17 for SQL Server};SERVER=server;DATABASE=db;UID=user;PWD=password"
FETCH_QUERY = (
    "SELECT timestamp, dia_traffic, cdn_traffic "
    "FROM traffic_table WHERE timestamp >= ? AND timestamp < ? ORDER BY timestamp"
)


def fetch_traffic_data(conn, start: dt.datetime, end: dt.datetime) -> pd.DataFrame:
    """Return traffic between start and end timestamps."""
    if pyodbc is None:
        raise RuntimeError("pyodbc is required for database access")
    return pd.read_sql(FETCH_QUERY, conn, params=(start, end))


def compute_baseline(
    df: pd.DataFrame,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Return resampled data, rolling mean and std over 7 days."""
    df = df.set_index("timestamp").sort_index()
    resampled = df.resample("15T").sum()
    baseline = resampled.rolling("7D").mean()
    std = resampled.rolling("7D").std()
    return resampled, baseline, std


def detect_anomaly(
    resampled: pd.DataFrame, baseline: pd.DataFrame, std: pd.DataFrame
) -> Tuple[bool, pd.Series, pd.Series]:
    """Determine if last quarter is an anomaly.

    Returns a tuple ``(is_anomaly, diff, mean)`` where ``diff`` is the
    difference between the last measurement and baseline mean, and ``mean``
    is the baseline mean for that quarter.
    """
    last = resampled.iloc[-1]
    # Exclude the most recent point when computing baseline statistics
    mean = baseline.iloc[-2]
    sigma = std.iloc[-2]
    diff = last - mean
    is_anomaly = (np.abs(diff) > 3 * sigma).any()
    return bool(is_anomaly), diff, mean


def plot_traffic(
    resampled: pd.DataFrame, baseline: pd.DataFrame, last_time: pd.Timestamp
) -> go.Figure:
    """Generate a plotly figure for the last 24h and baseline."""
    recent = resampled.loc[last_time - pd.Timedelta("24H") : last_time]
    recent_baseline = baseline.loc[recent.index]

    fig = go.Figure()
    for col, name in [("dia_traffic", "DIA"), ("cdn_traffic", "CDN")]:
        fig.add_trace(
            go.Scatter(x=recent.index, y=recent[col], mode="lines", name=name)
        )
        fig.add_trace(
            go.Scatter(
                x=recent_baseline.index,
                y=recent_baseline[col],
                mode="lines",
                name=f"{name} baseline",
                line=dict(dash="dash"),
            )
        )
        fig.add_trace(
            go.Scatter(
                x=[recent.index[-1]],
                y=[recent[col].iloc[-1]],
                mode="markers",
                name=f"{name} last quarter",
            )
        )

    fig.update_layout(
        title="DIA/CDN Traffic vs Baseline",
        xaxis_title="Time",
        yaxis_title="Traffic",
    )
    return fig


def send_email(fig: go.Figure, note: str, to_addr: str, subject: str) -> None:
    """Send an email with plotly HTML figure using Outlook."""
    if win32 is None:
        raise RuntimeError("win32com is required to send email via Outlook")

    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to_addr
    mail.Subject = subject
    mail.HTMLBody = f"<p>{note}</p>" + fig.to_html(include_plotlyjs="cdn")
    mail.Send()


def analyze(conn) -> None:
    """Fetch data, detect anomalies, and send email if needed."""
    now = dt.datetime.now()
    end = now.replace(second=0, microsecond=0) - dt.timedelta(
        minutes=(now.minute % 15) + 2
    )
    start = end - dt.timedelta(days=7)
    df = fetch_traffic_data(conn, start, end)
    resampled, baseline, std = compute_baseline(df)
    anomaly, diff, mean = detect_anomaly(resampled, baseline, std)
    if anomaly:
        notes = []
        for col, name in [("dia_traffic", "DIA"), ("cdn_traffic", "CDN")]:
            change = diff[col]
            direction = "increase" if change > 0 else "decrease"
            pct = (change / mean[col]) * 100 if mean[col] else 0
            notes.append(f"{name} {direction} of {abs(pct):.2f}%")
        note_text = " and ".join(notes) + " during the last quarter."
        fig = plot_traffic(resampled, baseline, resampled.index[-1])
        send_email(fig, note_text, "recipient@example.com", "Traffic anomaly detected")


def wait_loop() -> None:
    """Run analysis 2 minutes after each quarter hour."""
    if pyodbc is None:
        raise RuntimeError("pyodbc is required for database access")
    conn = pyodbc.connect(DB_CONNECTION_STRING)
    try:
        while True:
            now = dt.datetime.now()
            next_quarter = now.replace(
                minute=0, second=0, microsecond=0
            ) + dt.timedelta(minutes=15 * (now.minute // 15 + 1))
            run_time = next_quarter + dt.timedelta(minutes=2)
            time.sleep((run_time - now).total_seconds())
            analyze(conn)
    finally:
        conn.close()


if __name__ == "__main__":  # pragma: no cover
    wait_loop()
