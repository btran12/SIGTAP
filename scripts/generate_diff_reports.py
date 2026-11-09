#!/usr/bin/env python3
"""Generate HTML diff reports comparing consecutive SIGTAP unified tables.

The SIGTAP archives contain fixed-width text files described by ``layout.txt``.
Each record ends with a competence date (``DT_COMPETENCIA``) that changes every
month, so date-related columns are ignored when comparing records.

The script writes one HTML report per consecutive pair of monthly archives plus
an ``index.html`` that links them all together.

Usage:
    python3 scripts/generate_diff_reports.py [--tables-dir tabelas] [--output-dir diffs]
"""

from __future__ import annotations

import argparse
import hashlib
import html
import os
import re
import sys
import zipfile
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path

BINARY_EXTENSIONS = {
    ".xls",
    ".xlsx",
    ".pdf",
    ".zip",
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".ico",
}

MAX_DISPLAY_ROWS = 100
MAX_VALUE_LENGTH = 500

MONTHS = {
    "01": "January",
    "02": "February",
    "03": "March",
    "04": "April",
    "05": "May",
    "06": "June",
    "07": "July",
    "08": "August",
    "09": "September",
    "10": "October",
    "11": "November",
    "12": "December",
}

# Date formats that may appear inside free-text fields. Competence columns are
# excluded separately by name, so plain YYYYMM values are intentionally not
# matched here (they could be part of a code or value).
DATE_PATTERNS = [
    re.compile(r"\b\d{4}-\d{2}-\d{2}\b"),
    re.compile(r"\b\d{2}/\d{2}/\d{4}\b"),
    re.compile(r"\b\d{2}-\d{2}-\d{4}\b"),
    re.compile(r"\b\d{4}/\d{2}/\d{2}\b"),
    re.compile(r"\b\d{2}\.\d{2}\.\d{4}\b"),
]

PERIOD_RE = re.compile(r"TabelaUnificada_(\d{6})_")

CSS = """
:root {
  --bg: #f6f8fa;
  --panel: #ffffff;
  --border: #d0d7de;
  --text: #1f2328;
  --muted: #59636e;
  --accent: #0969da;
  --added-bg: #e6ffec;
  --added-text: #1a7f37;
  --removed-bg: #ffebe9;
  --removed-text: #cf222e;
  --modified-bg: #fff8c5;
  --modified-text: #7d4e00;
}
* { box-sizing: border-box; }
body {
  margin: 0;
  padding: 0 0 4rem;
  background: var(--bg);
  color: var(--text);
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  font-size: 15px;
  line-height: 1.5;
}
a { color: var(--accent); text-decoration: none; }
a:hover { text-decoration: underline; }
code {
  font-family: ui-monospace, SFMono-Regular, "SF Mono", Menlo, Consolas, monospace;
  font-size: 0.9em;
  background: rgba(175, 184, 193, 0.2);
  padding: 0.1em 0.35em;
  border-radius: 4px;
}
header.page-header {
  background: var(--panel);
  border-bottom: 1px solid var(--border);
  padding: 1.5rem max(1rem, calc((100vw - 1200px) / 2));
  margin-bottom: 1.5rem;
}
header.page-header h1 { margin: 0.35rem 0; font-size: 1.6rem; }
header.page-header p { margin: 0.25rem 0; color: var(--muted); }
header.page-header .back { font-weight: 600; }
.pager { display: flex; gap: 1rem; flex-wrap: wrap; margin-top: 0.75rem; font-weight: 600; }
.pager .spacer { flex: 1; }
main { max-width: 1200px; margin: 0 auto; padding: 0 1rem; }
.cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 0.75rem; margin-bottom: 1.75rem; }
.card { background: var(--panel); border: 1px solid var(--border); border-radius: 8px; padding: 0.9rem 1rem; }
.card .value { font-size: 1.6rem; font-weight: 700; line-height: 1.2; }
.card .label { color: var(--muted); font-size: 0.85rem; text-transform: uppercase; letter-spacing: 0.03em; }
.card.added .value { color: var(--added-text); }
.card.removed .value { color: var(--removed-text); }
.card.modified .value { color: var(--modified-text); }
section { margin-bottom: 2rem; }
h2 { font-size: 1.2rem; margin: 0 0 0.75rem; }
input.filter {
  width: 100%;
  max-width: 380px;
  padding: 0.45rem 0.65rem;
  margin-bottom: 0.75rem;
  border: 1px solid var(--border);
  border-radius: 6px;
  font-size: 0.95rem;
}
table { width: 100%; border-collapse: collapse; background: var(--panel); }
.table-wrap { overflow-x: auto; border: 1px solid var(--border); border-radius: 8px; margin-bottom: 0.75rem; }
th, td { border-bottom: 1px solid var(--border); padding: 0.45rem 0.6rem; text-align: left; vertical-align: top; }
th { background: #eef1f4; font-size: 0.8rem; text-transform: uppercase; letter-spacing: 0.03em; color: var(--muted); position: sticky; top: 0; }
tr:last-child td { border-bottom: none; }
tbody tr:hover td { background: #f6f8fa; }
.badge { display: inline-block; padding: 0.05rem 0.45rem; border-radius: 999px; font-size: 0.75rem; font-weight: 600; white-space: nowrap; }
.badge.added { background: var(--added-bg); color: var(--added-text); }
.badge.removed { background: var(--removed-bg); color: var(--removed-text); }
.badge.modified { background: var(--modified-bg); color: var(--modified-text); }
.badge.unchanged { background: #eaeef2; color: var(--muted); }
.badge.binary { background: #ddf4ff; color: #0969da; }
td.num { text-align: right; font-variant-numeric: tabular-nums; }
td.added { color: var(--added-text); font-weight: 600; }
td.removed { color: var(--removed-text); font-weight: 600; }
td.modified { color: var(--modified-text); font-weight: 600; }
details.file { background: var(--panel); border: 1px solid var(--border); border-radius: 8px; margin-bottom: 0.65rem; }
details.file > summary { cursor: pointer; padding: 0.65rem 0.85rem; font-weight: 600; list-style: none; display: flex; gap: 0.6rem; align-items: center; flex-wrap: wrap; }
details.file > summary::-webkit-details-marker { display: none; }
details.file > summary::before { content: "\\25B8"; color: var(--muted); font-size: 0.9em; }
details.file[open] > summary::before { content: "\\25BE"; }
details.file .body { padding: 0.25rem 0.85rem 0.85rem; border-top: 1px solid var(--border); }
details.file h3 { font-size: 0.95rem; margin: 0.85rem 0 0.4rem; }
.counts { color: var(--muted); font-weight: 400; font-size: 0.85rem; }
.note { color: var(--muted); font-size: 0.85rem; margin: 0.25rem 0 0.5rem; }
.empty { color: var(--muted); font-style: italic; }
.legend { display: flex; gap: 1rem; flex-wrap: wrap; margin: 0.5rem 0 1rem; }
.legend span { display: inline-flex; align-items: center; gap: 0.35rem; color: var(--muted); font-size: 0.9rem; }
footer { max-width: 1200px; margin: 2rem auto 0; padding: 1rem; color: var(--muted); font-size: 0.85rem; text-align: center; }
.toast-container {
  position: fixed;
  right: 1rem;
  bottom: 1rem;
  z-index: 1000;
  display: flex;
  flex-direction: column;
  gap: 0.5rem;
  width: min(380px, calc(100vw - 2rem));
}
.toast {
  background: var(--panel);
  border: 1px solid var(--border);
  border-left: 4px solid var(--accent);
  border-radius: 8px;
  box-shadow: 0 8px 24px rgba(31, 35, 40, 0.18);
  padding: 0.75rem 0.9rem;
  display: flex;
  gap: 0.6rem;
  align-items: flex-start;
  animation: toast-in 0.18s ease-out;
}
.toast.success { border-left-color: var(--added-text); }
.toast.warning { border-left-color: var(--modified-text); }
.toast.error { border-left-color: var(--removed-text); }
.toast.toast-out { opacity: 0; transform: translateY(6px); transition: opacity 0.16s ease-in, transform 0.16s ease-in; }
.toast .toast-body { flex: 1; min-width: 0; }
.toast .toast-title { font-weight: 700; margin: 0 0 0.15rem; }
.toast .toast-message { margin: 0; color: var(--muted); font-size: 0.9rem; overflow-wrap: anywhere; }
.toast .toast-action { display: inline-block; margin-top: 0.35rem; font-weight: 600; }
.toast .toast-close {
  border: 0;
  background: transparent;
  color: var(--muted);
  font-size: 1.15rem;
  line-height: 1;
  cursor: pointer;
  padding: 0.1rem 0.2rem;
}
.toast .toast-close:hover { color: var(--text); }
.notify-toggle {
  border: 1px solid var(--border);
  background: var(--panel);
  color: var(--text);
  border-radius: 999px;
  padding: 0.35rem 0.75rem;
  font: inherit;
  font-size: 0.85rem;
  font-weight: 600;
  cursor: pointer;
}
.notify-toggle:hover { border-color: var(--accent); color: var(--accent); }
.notify-toggle[disabled] { opacity: 0.6; cursor: not-allowed; }
@keyframes toast-in {
  from { opacity: 0; transform: translateY(8px); }
  to { opacity: 1; transform: translateY(0); }
}
@media (max-width: 600px) {
  .toast-container { left: 1rem; right: 1rem; bottom: 1rem; width: auto; }
}
"""

TOAST_SCRIPT = r"""
<script>
(function () {
  "use strict";

  var SEEN_KEY = "sigtap.lastSeenPeriod";
  var ENABLED_KEY = "sigtap.notificationsEnabled";

  var body = document.body;
  var latestPeriod = body.getAttribute("data-latest-period") || "";
  var latestLabel = body.getAttribute("data-latest-label") || latestPeriod;
  var latestHref = body.getAttribute("data-latest-href") || "";
  var latestDiffLabel = body.getAttribute("data-latest-diff-label") || "";

  function read(key) {
    try { return window.localStorage.getItem(key); } catch (error) { return null; }
  }
  function write(key, value) {
    try { window.localStorage.setItem(key, value); } catch (error) { /* ignore */ }
  }

  function toastContainer() {
    var element = document.getElementById("toast-container");
    if (!element) {
      element = document.createElement("div");
      element.id = "toast-container";
      element.className = "toast-container";
      element.setAttribute("role", "status");
      element.setAttribute("aria-live", "polite");
      document.body.appendChild(element);
    }
    return element;
  }

  function dismissToast(element) {
    if (!element || element.classList.contains("toast-out")) return;
    element.classList.add("toast-out");
    window.setTimeout(function () {
      if (element.parentNode) element.parentNode.removeChild(element);
    }, 180);
  }

  function showToast(options) {
    options = options || {};
    var element = document.createElement("div");
    element.className = "toast " + (options.type || "info");

    var content = document.createElement("div");
    content.className = "toast-body";

    var title = document.createElement("p");
    title.className = "toast-title";
    title.textContent = options.title || "";
    content.appendChild(title);

    if (options.message) {
      var message = document.createElement("p");
      message.className = "toast-message";
      message.textContent = options.message;
      content.appendChild(message);
    }

    if (options.actionHref) {
      var action = document.createElement("a");
      action.className = "toast-action";
      action.href = options.actionHref;
      action.textContent = options.actionLabel || "View";
      content.appendChild(action);
    }

    var close = document.createElement("button");
    close.type = "button";
    close.className = "toast-close";
    close.setAttribute("aria-label", "Dismiss notification");
    close.textContent = "\u00d7";
    close.addEventListener("click", function () { dismissToast(element); });

    element.appendChild(content);
    element.appendChild(close);
    toastContainer().appendChild(element);

    var duration = options.duration == null ? 9000 : options.duration;
    if (duration > 0) {
      var timer = window.setTimeout(function () { dismissToast(element); }, duration);
      element.addEventListener("mouseenter", function () { window.clearTimeout(timer); });
      element.addEventListener("mouseleave", function () {
        timer = window.setTimeout(function () { dismissToast(element); }, 3000);
      });
    }
    return element;
  }

  window.sigtapShowToast = showToast;

  function notificationsSupported() {
    return typeof window.Notification !== "undefined" && window.isSecureContext !== false;
  }

  function notificationsEnabled() {
    return read(ENABLED_KEY) === "1" &&
      notificationsSupported() &&
      window.Notification.permission === "granted";
  }

  function updateToggle() {
    var button = document.getElementById("notify-toggle");
    if (!button) return;
    if (!notificationsSupported()) {
      button.hidden = true;
      return;
    }
    if (window.Notification.permission === "denied") {
      button.textContent = "\ud83d\udd15 Notifications blocked";
      button.disabled = true;
      button.title = "Notifications are blocked in your browser settings.";
      return;
    }
    if (notificationsEnabled()) {
      button.textContent = "\ud83d\udd15 Disable notifications";
      button.title = "Stop on-device update notifications.";
    } else {
      button.textContent = "\ud83d\udd14 Enable notifications";
      button.title = "Get an on-device notification when a new monthly table is published.";
    }
  }

  function notifyDevice() {
    if (!notificationsEnabled()) return;
    try {
      var notification = new window.Notification("New SIGTAP update", {
        body: latestDiffLabel
          ? latestLabel + " is available. " + latestDiffLabel + "."
          : latestLabel + " is available.",
        tag: "sigtap-update-" + latestPeriod
      });
      notification.onclick = function () {
        window.focus();
        if (latestHref) window.location.href = latestHref;
        notification.close();
      };
    } catch (error) { /* notifications are best-effort */ }
  }

  function checkForUpdates() {
    if (!latestPeriod) return;
    var seen = read(SEEN_KEY);
    if (!seen) {
      write(SEEN_KEY, latestPeriod);
      return;
    }
    if (String(seen) >= String(latestPeriod)) return;
    write(SEEN_KEY, latestPeriod);
    showToast({
      title: "New SIGTAP update",
      message: latestDiffLabel
        ? latestLabel + " is available. " + latestDiffLabel + "."
        : latestLabel + " is available.",
      type: "success",
      actionHref: latestHref || null,
      actionLabel: latestDiffLabel ? "View latest comparison" : "View downloads",
      duration: 12000
    });
    notifyDevice();
  }

  function setupToggle() {
    var button = document.getElementById("notify-toggle");
    if (!button) return;
    button.addEventListener("click", function () {
      if (!notificationsSupported()) {
        showToast({
          title: "Notifications unavailable",
          message: "This browser does not support on-device notifications.",
          type: "warning"
        });
        return;
      }
      if (window.Notification.permission === "denied") {
        showToast({
          title: "Notifications blocked",
          message: "Allow notifications for this site in your browser settings.",
          type: "warning"
        });
        return;
      }
      if (notificationsEnabled()) {
        write(ENABLED_KEY, "0");
        updateToggle();
        showToast({
          title: "Notifications disabled",
          message: "You will no longer receive on-device update alerts.",
          type: "info"
        });
        return;
      }

      var finished = false;
      function finish() {
        if (finished) return;
        finished = true;
        if (window.Notification.permission === "granted") {
          write(ENABLED_KEY, "1");
          updateToggle();
          showToast({
            title: "Notifications enabled",
            message: "You will be notified when a new monthly table is published.",
            type: "success"
          });
        } else {
          updateToggle();
          showToast({
            title: "Notifications not enabled",
            message: "Permission was not granted.",
            type: "warning"
          });
        }
      }

      if (window.Notification.permission === "granted") {
        finish();
        return;
      }
      var result = window.Notification.requestPermission(finish);
      if (result && typeof result.then === "function") {
        result.then(finish);
      }
    });
  }

  document.addEventListener("DOMContentLoaded", function () {
    updateToggle();
    setupToggle();
    checkForUpdates();
  });
})();
</script>
"""


def strip_dates(value: str) -> str:
    """Replace date-like tokens so date-only changes are ignored."""
    for pattern in DATE_PATTERNS:
        value = pattern.sub("<date>", value)
    return value


def shorten(value: str) -> str:
    if len(value) > MAX_VALUE_LENGTH:
        return value[:MAX_VALUE_LENGTH] + "\u2026"
    return value


def escape(value: str) -> str:
    return html.escape(shorten(value), quote=True)


def period_from_name(name: str) -> str | None:
    match = PERIOD_RE.search(name)
    return match.group(1) if match else None


def period_label(period: str) -> str:
    if len(period) == 6 and period[:4].isdigit() and period[4:] in MONTHS:
        return f"{MONTHS[period[4:]]}/{period[:4]}"
    return period


def human_size(num_bytes: int) -> str:
    size = float(num_bytes)
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024 or unit == "GB":
            if unit == "B":
                return f"{int(size)} B"
            return f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} TB"


def relative_url(from_dir: Path, target: Path) -> str:
    """Return a POSIX-style relative URL from one path to another."""
    try:
        relative = os.path.relpath(target, start=from_dir)
    except ValueError:
        relative = str(target)
    return Path(relative).as_posix()


def body_attributes(latest: dict, latest_href: str) -> str:
    """Data attributes consumed by the toast/notification script."""
    return (
        f'data-latest-period="{html.escape(latest.get("period", ""))}" '
        f'data-latest-label="{html.escape(latest.get("label", ""))}" '
        f'data-latest-href="{html.escape(latest_href)}" '
        f'data-latest-diff-label="{html.escape(latest.get("diff_label", ""))}"'
    )


def is_date_column(name: str) -> bool:
    upper = name.upper()
    return (
        upper.startswith("DT")
        or upper.endswith("_DT")
        or "COMPETENCIA" in upper
        or "DATA" in upper
    )


def is_key_column(name: str) -> bool:
    upper = name.upper()
    return upper.startswith("CO_") or upper.startswith("NU_")


def parse_layout_text(text: str) -> dict[str, list[tuple[str, int, int]]]:
    """Parse the combined ``layout.txt`` (table name followed by CSV header/rows)."""
    layouts: dict[str, list[tuple[str, int, int]]] = {}
    current: str | None = None
    for raw in text.splitlines():
        line = raw.strip().lstrip("\ufeff")
        if not line:
            continue
        if line.lower().startswith("coluna,"):
            continue
        if "," not in line:
            current = line
            layouts.setdefault(current, [])
            continue
        if current is None:
            continue
        parts = [part.strip() for part in line.split(",")]
        if len(parts) < 5:
            continue
        try:
            start, end = int(parts[2]), int(parts[3])
        except ValueError:
            continue
        layouts[current].append((parts[0], start, end))
    return layouts


def parse_single_layout(text: str) -> list[tuple[str, int, int]]:
    """Parse a per-table ``*_layout.txt`` file (CSV header/rows only)."""
    columns: list[tuple[str, int, int]] = []
    for raw in text.splitlines():
        line = raw.strip().lstrip("\ufeff")
        if not line or line.lower().startswith("coluna,"):
            continue
        parts = [part.strip() for part in line.split(",")]
        if len(parts) < 5:
            continue
        try:
            start, end = int(parts[2]), int(parts[3])
        except ValueError:
            continue
        columns.append((parts[0], start, end))
    return columns


def layout_for(
    name: str,
    layouts: dict[str, list[tuple[str, int, int]]],
    archive: zipfile.ZipFile,
) -> list[tuple[str, int, int]] | None:
    stem = Path(name).stem
    if layouts.get(stem):
        return layouts[stem]
    per_file = f"{stem}_layout.txt"
    if per_file in archive.namelist():
        return parse_single_layout(archive.read(per_file).decode("latin-1"))
    return None


def is_binary(name: str, data: bytes) -> bool:
    if Path(name).suffix.lower() in BINARY_EXTENSIONS:
        return True
    return b"\x00" in data[:8192]


def parse_rows(text: str, columns: list[tuple[str, int, int]]) -> list[tuple[str, ...]]:
    rows: list[tuple[str, ...]] = []
    for line in text.splitlines():
        if not line.strip():
            continue
        rows.append(tuple(line[start - 1 : end].strip() for _, start, end in columns))
    return rows


def compare_layout(
    old_text: str,
    new_text: str,
    layout: list[tuple[str, int, int]],
) -> dict:
    columns = [column for column in layout if not is_date_column(column[0])]
    key_positions = [index for index, column in enumerate(columns) if is_key_column(column[0])]
    attr_positions = [index for index, column in enumerate(columns) if not is_key_column(column[0])]

    old_rows = parse_rows(old_text, columns)
    new_rows = parse_rows(new_text, columns)

    def key_of(row: tuple[str, ...]) -> tuple[str, ...]:
        if key_positions:
            return tuple(row[index] for index in key_positions)
        return row

    old_map: dict[tuple[str, ...], list[tuple[str, ...]]] = defaultdict(list)
    new_map: dict[tuple[str, ...], list[tuple[str, ...]]] = defaultdict(list)
    for row in old_rows:
        old_map[key_of(row)].append(row)
    for row in new_rows:
        new_map[key_of(row)].append(row)

    added: list[tuple[str, ...]] = []
    removed: list[tuple[str, ...]] = []
    modified: list[tuple[tuple[str, ...], list[tuple[str, str, str]]]] = []

    for key in sorted(set(old_map) | set(new_map), key=lambda item: tuple(map(str, item))):
        old_group = sorted(old_map.get(key, []))
        new_group = sorted(new_map.get(key, []))
        common = min(len(old_group), len(new_group))
        for index in range(common):
            changes = []
            for position in attr_positions:
                if strip_dates(old_group[index][position]) != strip_dates(new_group[index][position]):
                    changes.append(
                        (
                            columns[position][0],
                            old_group[index][position],
                            new_group[index][position],
                        )
                    )
            if changes:
                modified.append((key, changes))
        removed.extend(old_group[common:])
        added.extend(new_group[common:])

    return {
        "kind": "layout",
        "columns": [column[0] for column in columns],
        "key_columns": [columns[position][0] for position in key_positions],
        "added": added,
        "removed": removed,
        "modified": modified,
    }


def compare_lines(old_text: str, new_text: str) -> dict:
    def prepare(text: str) -> tuple[Counter, dict[str, str]]:
        counter: Counter = Counter()
        raw_by_normalized: dict[str, str] = {}
        for raw in text.splitlines():
            line = raw.rstrip()
            if not line.strip():
                continue
            normalized = strip_dates(line)
            counter[normalized] += 1
            raw_by_normalized.setdefault(normalized, line)
        return counter, raw_by_normalized

    old_counter, old_raw = prepare(old_text)
    new_counter, new_raw = prepare(new_text)

    added_norm = new_counter - old_counter
    removed_norm = old_counter - new_counter

    added = [new_raw.get(norm, norm) for norm, count in added_norm.items() for _ in range(count)]
    removed = [old_raw.get(norm, norm) for norm, count in removed_norm.items() for _ in range(count)]

    return {
        "kind": "lines",
        "columns": [],
        "key_columns": [],
        "added": added,
        "removed": removed,
        "modified": [],
    }


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def compare_archive_pair(
    previous_zip: Path,
    current_zip: Path,
) -> tuple[dict, list[dict]]:
    """Compare two archives and return (summary, file results)."""
    with zipfile.ZipFile(previous_zip) as old_zip, zipfile.ZipFile(current_zip) as new_zip:
        old_names = {info.filename for info in old_zip.infolist() if not info.is_dir()}
        new_names = {info.filename for info in new_zip.infolist() if not info.is_dir()}
        all_names = sorted(old_names | new_names)

        layouts = {}
        if "layout.txt" in new_zip.namelist():
            layouts = parse_layout_text(new_zip.read("layout.txt").decode("latin-1"))
        elif "layout.txt" in old_zip.namelist():
            layouts = parse_layout_text(old_zip.read("layout.txt").decode("latin-1"))

        results: list[dict] = []
        for name in all_names:
            old_data = old_zip.read(name) if name in old_names else None
            new_data = new_zip.read(name) if name in new_names else None

            if old_data is None:
                results.append(
                    {
                        "name": name,
                        "status": "added_file",
                        "added": 0,
                        "removed": 0,
                        "modified": 0,
                        "kind": "none",
                        "columns": [],
                        "key_columns": [],
                        "added_rows": [],
                        "removed_rows": [],
                        "modified_rows": [],
                    }
                )
                continue
            if new_data is None:
                results.append(
                    {
                        "name": name,
                        "status": "removed_file",
                        "added": 0,
                        "removed": 0,
                        "modified": 0,
                        "kind": "none",
                        "columns": [],
                        "key_columns": [],
                        "added_rows": [],
                        "removed_rows": [],
                        "modified_rows": [],
                    }
                )
                continue

            if is_binary(name, new_data):
                changed = sha256(old_data) != sha256(new_data)
                results.append(
                    {
                        "name": name,
                        "status": "changed" if changed else "unchanged",
                        "added": 0,
                        "removed": 0,
                        "modified": 0,
                        "kind": "binary",
                        "columns": [],
                        "key_columns": [],
                        "added_rows": [],
                        "removed_rows": [],
                        "modified_rows": [],
                    }
                )
                continue

            old_text = old_data.decode("latin-1")
            new_text = new_data.decode("latin-1")
            layout = layout_for(name, layouts, new_zip) or layout_for(name, layouts, old_zip)

            if layout:
                comparison = compare_layout(old_text, new_text, layout)
            else:
                comparison = compare_lines(old_text, new_text)

            added_count = len(comparison["added"])
            removed_count = len(comparison["removed"])
            modified_count = len(comparison["modified"])
            changed = bool(added_count or removed_count or modified_count)

            results.append(
                {
                    "name": name,
                    "status": "changed" if changed else "unchanged",
                    "added": added_count,
                    "removed": removed_count,
                    "modified": modified_count,
                    "kind": comparison["kind"],
                    "columns": comparison["columns"],
                    "key_columns": comparison["key_columns"],
                    "added_rows": comparison["added"],
                    "removed_rows": comparison["removed"],
                    "modified_rows": comparison["modified"],
                }
            )

    changed_files = sum(1 for result in results if result["status"] != "unchanged")
    summary = {
        "previous_zip": previous_zip.name,
        "current_zip": current_zip.name,
        "files_total": len(results),
        "files_changed": changed_files,
        "added": sum(result["added"] for result in results),
        "removed": sum(result["removed"] for result in results),
        "modified": sum(result["modified"] for result in results),
    }
    return summary, results


def render_added_removed_table(
    title: str,
    rows: list,
    columns: list[str],
    css_class: str,
) -> str:
    if not rows:
        return ""
    displayed = rows[:MAX_DISPLAY_ROWS]
    parts = [f'<h3 class="{css_class}">{html.escape(title)} ({len(rows)})</h3>']
    if len(rows) > MAX_DISPLAY_ROWS:
        parts.append(
            f'<p class="note">Showing the first {MAX_DISPLAY_ROWS} of {len(rows)} records.</p>'
        )
    parts.append('<div class="table-wrap"><table><thead><tr>')
    if columns:
        parts.extend(f"<th>{html.escape(column)}</th>" for column in columns)
    else:
        parts.append("<th>Line</th>")
    parts.append("</tr></thead><tbody>")
    for row in displayed:
        parts.append("<tr>")
        values = row if isinstance(row, tuple) else (row,)
        for value in values:
            parts.append(f"<td>{escape(str(value))}</td>")
        parts.append("</tr>")
    parts.append("</tbody></table></div>")
    return "".join(parts)


def render_modified_table(
    modified: list[tuple[tuple[str, ...], list[tuple[str, str, str]]]],
    key_columns: list[str],
) -> str:
    if not modified:
        return ""
    parts = [f'<h3 class="modified">Modified records ({len(modified)})</h3>']
    if len(modified) > MAX_DISPLAY_ROWS:
        parts.append(
            f'<p class="note">Showing the first {MAX_DISPLAY_ROWS} of {len(modified)} records.</p>'
        )
    parts.append(
        '<div class="table-wrap"><table><thead><tr>'
        "<th>Record</th><th>Field</th><th>Previous value</th><th>Current value</th>"
        "</tr></thead><tbody>"
    )
    for key, changes in modified[:MAX_DISPLAY_ROWS]:
        if key_columns:
            record = ", ".join(
                f"{column}={value}" for column, value in zip(key_columns, key)
            )
        else:
            record = "\u2014"
        for field, old_value, new_value in changes:
            parts.append(
                f"<tr><td><code>{escape(record)}</code></td>"
                f"<td>{html.escape(field)}</td>"
                f'<td class="removed">{escape(old_value)}</td>'
                f'<td class="added">{escape(new_value)}</td></tr>'
            )
    parts.append("</tbody></table></div>")
    return "".join(parts)


def render_file_details(result: dict) -> str:
    status = result["status"]
    if status == "added_file":
        body = '<p class="empty">File added in the current table.</p>'
    elif status == "removed_file":
        body = '<p class="empty">File removed from the current table.</p>'
    elif result["kind"] == "binary":
        if status == "changed":
            body = '<p class="empty">Binary file changed (SHA-256 differs).</p>'
        else:
            body = '<p class="empty">Binary file unchanged.</p>'
    else:
        body_parts = []
        body_parts.append(
            render_added_removed_table(
                "Added records",
                result["added_rows"],
                result["columns"],
                "added",
            )
        )
        body_parts.append(
            render_added_removed_table(
                "Removed records",
                result["removed_rows"],
                result["columns"],
                "removed",
            )
        )
        body_parts.append(
            render_modified_table(result["modified_rows"], result["key_columns"])
        )
        body = "".join(body_parts) or '<p class="empty">No differences.</p>'

    counts = []
    if result["added"]:
        counts.append(f'<span class="badge added">+{result["added"]}</span>')
    if result["removed"]:
        counts.append(f'<span class="badge removed">-{result["removed"]}</span>')
    if result["modified"]:
        counts.append(f'<span class="badge modified">~{result["modified"]}</span>')
    if result["kind"] == "binary" and status == "changed":
        counts.append('<span class="badge binary">binary</span>')
    if not counts and status != "added_file" and status != "removed_file":
        counts.append('<span class="badge unchanged">unchanged</span>')

    status_badge = ""
    if status == "added_file":
        status_badge = '<span class="badge added">file added</span>'
    elif status == "removed_file":
        status_badge = '<span class="badge removed">file removed</span>'

    return (
        f'<details class="file" id="file-{html.escape(result["name"].replace(".", "-"))}">'
        f"<summary>{status_badge}<code>{html.escape(result['name'])}</code>"
        f'<span class="counts">{" ".join(counts)}</span></summary>'
        f'<div class="body">{body}</div>'
        "</details>"
    )


def render_file_summary_table(results: list[dict]) -> str:
    rows = []
    for result in results:
        status = result["status"]
        if status == "added_file":
            badge = '<span class="badge added">file added</span>'
        elif status == "removed_file":
            badge = '<span class="badge removed">file removed</span>'
        elif result["kind"] == "binary":
            badge = (
                '<span class="badge binary">binary changed</span>'
                if status == "changed"
                else '<span class="badge unchanged">unchanged</span>'
            )
        elif status == "changed":
            badge = '<span class="badge modified">changed</span>'
        else:
            badge = '<span class="badge unchanged">unchanged</span>'

        anchor = result["name"].replace(".", "-")
        if status == "unchanged":
            name_cell = f'<code>{html.escape(result["name"])}</code>'
        else:
            name_cell = (
                f'<a href="#file-{html.escape(anchor)}">'
                f'<code>{html.escape(result["name"])}</code></a>'
            )
        rows.append(
            "<tr>"
            f"<td>{name_cell}</td>"
            f"<td>{badge}</td>"
            f'<td class="num added">{result["added"] or ""}</td>'
            f'<td class="num removed">{result["removed"] or ""}</td>'
            f'<td class="num modified">{result["modified"] or ""}</td>'
            "</tr>"
        )
    return (
        '<div class="table-wrap"><table id="files-summary"><thead><tr>'
        "<th>File</th><th>Status</th><th>Added</th><th>Removed</th><th>Modified</th>"
        "</tr></thead><tbody>"
        + "".join(rows)
        + "</tbody></table></div>"
    )


def render_report(
    summary: dict,
    results: list[dict],
    previous_period: str,
    current_period: str,
    previous_link: str | None,
    next_link: str | None,
    generated_at: str,
    home_link: str,
    latest: dict,
) -> str:
    previous_label = period_label(previous_period)
    current_label = period_label(current_period)

    pager_parts = [
        f'<a href="{home_link}">\u2190 Home</a>',
        f'<a href="{home_link}#downloads">\u2b07\ufe0f Downloads</a>',
        '<a href="index.html">All reports</a>',
        '<button id="notify-toggle" class="notify-toggle" type="button">\U0001F514 Enable notifications</button>',
    ]
    pager_parts.append('<span class="spacer"></span>')
    if previous_link:
        pager_parts.append(f'<a href="{previous_link}">\u2190 Previous month</a>')
    if next_link:
        pager_parts.append(f'<a href="{next_link}">Next month \u2192</a>')

    file_details = "".join(
        render_file_details(result)
        for result in results
        if result["status"] != "unchanged"
    )
    if not file_details:
        file_details = '<p class="empty">No differences found in any file.</p>'

    previous_zip = html.escape(summary['previous_zip'])
    current_zip = html.escape(summary['current_zip'])
    files_changed = summary['files_changed']
    added = summary['added']
    removed = summary['removed']
    modified = summary['modified']
    files_total = summary['files_total']
    pager = ''.join(pager_parts)
    latest_href = latest["filename"] if latest.get("filename") else f"{home_link}#downloads"
    body_attrs = body_attributes(latest, latest_href)

    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SIGTAP diff: {html.escape(previous_label)} \u2192 {html.escape(current_label)}</title>
<style>{CSS}</style>
</head>
<body {body_attrs}>
<header class="page-header">
  <a class="back" href="index.html">\u2190 All reports</a>
  <h1>SIGTAP Diff Report</h1>
  <p><strong>{html.escape(previous_label)} \u2192 {html.escape(current_label)}</strong></p>
  <p>Previous: <code>{previous_zip}</code><br>
     Current: <code>{current_zip}</code></p>
  <p>Generated {html.escape(generated_at)} \u00b7 Date-related changes (competence columns and date values) are ignored.</p>
  <nav class="pager">{pager}</nav>
</header>
<main>
  <section class="cards">
    <div class="card"><div class="value">{files_changed}</div><div class="label">Files changed</div></div>
    <div class="card added"><div class="value">{added}</div><div class="label">Added records</div></div>
    <div class="card removed"><div class="value">{removed}</div><div class="label">Removed records</div></div>
    <div class="card modified"><div class="value">{modified}</div><div class="label">Modified records</div></div>
    <div class="card"><div class="value">{files_total}</div><div class="label">Files compared</div></div>
  </section>

  <section>
    <h2>Files</h2>
    <input class="filter" id="file-filter" type="search" placeholder="Filter files\u2026" aria-label="Filter files">
    {render_file_summary_table(results)}
  </section>

  <section>
    <h2>Details</h2>
    <div class="legend">
      <span><span class="badge added">+</span> added</span>
      <span><span class="badge removed">-</span> removed</span>
      <span><span class="badge modified">~</span> modified</span>
    </div>
    <div id="file-details">{file_details}</div>
  </section>
</main>
<footer>Generated automatically from the DATASUS SIGTAP unified tables.</footer>
<script>
(function () {{
  var input = document.getElementById("file-filter");
  if (!input) return;
  input.addEventListener("input", function () {{
    var term = input.value.trim().toLowerCase();
    document.querySelectorAll("#files-summary tbody tr").forEach(function (row) {{
      row.style.display = row.textContent.toLowerCase().indexOf(term) !== -1 ? "" : "none";
    }});
    document.querySelectorAll("#file-details details.file").forEach(function (item) {{
      var name = item.querySelector("code");
      var text = name ? name.textContent.toLowerCase() : "";
      item.style.display = text.indexOf(term) !== -1 ? "" : "none";
    }});
  }});
}})();
</script>
{TOAST_SCRIPT}
</body>
</html>
"""


def render_reports_table(reports: list[dict], link_prefix: str = "") -> str:
    rows = []
    for report in sorted(reports, key=lambda item: item["current_period"], reverse=True):
        href = f"{link_prefix}{report['filename']}"
        rows.append(
            "<tr>"
            f'<td><a href="{html.escape(href)}">'
            f'{html.escape(period_label(report["previous_period"]))} \u2192 '
            f'{html.escape(period_label(report["current_period"]))}</a></td>'
            f'<td><code>{html.escape(report["previous_zip"])}</code></td>'
            f'<td><code>{html.escape(report["current_zip"])}</code></td>'
            f'<td class="num">{report["files_changed"]}</td>'
            f'<td class="num added">{report["added"]}</td>'
            f'<td class="num removed">{report["removed"]}</td>'
            f'<td class="num modified">{report["modified"]}</td>'
            "</tr>"
        )

    if not rows:
        return '<p class="empty">No monthly comparisons available yet.</p>'

    return (
        '<div class="table-wrap"><table><thead><tr>'
        "<th>Comparison</th><th>Previous archive</th><th>Current archive</th>"
        "<th>Files changed</th><th>Added</th><th>Removed</th><th>Modified</th>"
        "</tr></thead><tbody>"
        + "".join(rows)
        + "</tbody></table></div>"
    )


def render_index(
    reports: list[dict],
    generated_at: str,
    home_link: str,
    latest: dict,
) -> str:
    table = render_reports_table(reports)
    latest_href = latest["filename"] if latest.get("filename") else f"{home_link}#downloads"
    body_attrs = body_attributes(latest, latest_href)

    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SIGTAP Diff Reports</title>
<style>{CSS}</style>
</head>
<body {body_attrs}>
<header class="page-header">
  <a class="back" href="{home_link}">\u2190 Home</a>
  <h1>SIGTAP Diff Reports</h1>
  <p>Monthly comparisons between consecutive DATASUS SIGTAP unified tables.</p>
  <p>Generated {html.escape(generated_at)} \u00b7 Date-related changes are ignored.</p>
  <nav class="pager">
    <a href="{home_link}#downloads">\u2b07\ufe0f Downloads</a>
    <a href="{home_link}">Home</a>
    <button id="notify-toggle" class="notify-toggle" type="button">\U0001F514 Enable notifications</button>
  </nav>
</header>
<main>
  <section>
    <h2>Monthly comparisons</h2>
    {table}
  </section>
  <section>
    <h2>About</h2>
    <p>Each report compares every file in the previous archive against the current one and lists added,
       removed and modified records. Competence-date columns (such as <code>DT_COMPETENCIA</code>) and
       date-like values are ignored so that the monthly competence rollover does not appear as a change.</p>
  </section>
</main>
<footer>Generated automatically from the DATASUS SIGTAP unified tables.</footer>
{TOAST_SCRIPT}
</body>
</html>
"""


def render_home(
    archives: list[dict],
    reports: list[dict],
    generated_at: str,
    tables_link: str,
    diffs_link: str,
    notes_link: str,
    notes_count: int,
    latest: dict,
) -> str:
    total_size = sum(item["size_bytes"] for item in archives)
    latest_href = f"{diffs_link}/{latest['filename']}" if latest.get("filename") else "#downloads"
    body_attrs = body_attributes(latest, latest_href)
    table_rows = []
    for item in sorted(archives, key=lambda entry: entry["period"], reverse=True):
        href = f"{tables_link}/{item['name']}"
        table_rows.append(
            "<tr>"
            f'<td><code>{html.escape(item["name"])}</code></td>'
            f'<td>{html.escape(period_label(item["period"]))}</td>'
            f'<td class="num">{html.escape(item["size"])}</td>'
            f'<td><a href="{html.escape(href)}">\u2b07\ufe0f Download</a></td>'
            "</tr>"
        )

    if table_rows:
        tables_table = (
            '<div class="table-wrap"><table><thead><tr>'
            "<th>File</th><th>Period</th><th>Size</th><th>Download</th>"
            "</tr></thead><tbody>"
            + "".join(table_rows)
            + "</tbody></table></div>"
        )
    else:
        tables_table = '<p class="empty">No tables available yet.</p>'

    reports_table = render_reports_table(reports, f"{diffs_link}/")

    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>DATASUS SIGTAP Unified Table</title>
<style>{CSS}</style>
</head>
<body {body_attrs}>
<header class="page-header">
  <h1>DATASUS SIGTAP Unified Table</h1>
  <p>Automated mirror of the DATASUS SIGTAP unified tables.</p>
  <p>Generated {html.escape(generated_at)} \u00b7 {len(archives)} tables \u00b7 {len(reports)} diff reports</p>
  <nav class="pager">
    <a href="#downloads">\u2b07\ufe0f Downloads</a>
    <a href="#diffs">Diff reports</a>
    <a href="{notes_link}">Technical notes</a>
    <button id="notify-toggle" class="notify-toggle" type="button">\U0001F514 Enable notifications</button>
  </nav>
</header>
<main>
  <section class="cards">
    <div class="card"><div class="value">{len(archives)}</div><div class="label">Monthly tables</div></div>
    <div class="card"><div class="value">{human_size(total_size)}</div><div class="label">Total size</div></div>
    <div class="card"><div class="value">{len(reports)}</div><div class="label">Diff reports</div></div>
    <div class="card"><div class="value">{notes_count}</div><div class="label">Technical notes</div></div>
  </section>

  <section id="downloads">
    <h2>Unified Tables</h2>
    <p class="note">Download the most recent monthly archives from DATASUS.</p>
    {tables_table}
  </section>

  <section id="diffs">
    <h2>Monthly Diff Reports</h2>
    <p class="note">Compare consecutive monthly tables. Date-related changes are ignored.</p>
    {reports_table}
  </section>

  <section id="notes">
    <h2>Technical Notes</h2>
    <p><a href="{notes_link}">Browse the technical notes folder</a> ({notes_count} documents).</p>
  </section>
</main>
<footer>Generated automatically from the DATASUS SIGTAP unified tables.</footer>
{TOAST_SCRIPT}
</body>
</html>
"""


def discover_archives(tables_dir: Path) -> list[tuple[str, Path]]:
    archives: dict[str, Path] = {}
    for path in sorted(tables_dir.glob("TabelaUnificada_*.zip")):
        period = period_from_name(path.name)
        if period:
            archives[period] = path
    return sorted(archives.items())


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--tables-dir", default="tabelas", help="Directory containing the ZIP archives")
    parser.add_argument("--output-dir", default="diffs", help="Directory for the generated diff HTML files")
    parser.add_argument("--index-file", default="index.html", help="Path of the root landing page")
    parser.add_argument("--notes-dir", default="notastecnicas", help="Directory containing the technical notes")
    args = parser.parse_args(argv)

    tables_dir = Path(args.tables_dir)
    output_dir = Path(args.output_dir)
    index_file = Path(args.index_file)
    notes_dir = Path(args.notes_dir)

    archives = discover_archives(tables_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    index_file.parent.mkdir(parents=True, exist_ok=True)
    for old_report in output_dir.glob("*.html"):
        old_report.unlink()

    generated_at = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")

    index_dir = index_file.parent
    tables_link = relative_url(index_dir, tables_dir)
    diffs_link = relative_url(index_dir, output_dir)
    notes_link = relative_url(index_dir, notes_dir)
    home_from_diffs = relative_url(output_dir, index_file)
    notes_count = len(list(notes_dir.glob("*.pdf"))) if notes_dir.is_dir() else 0

    archives_info = [
        {
            "period": period,
            "name": path.name,
            "size": human_size(path.stat().st_size),
            "size_bytes": path.stat().st_size,
        }
        for period, path in archives
    ]

    latest_period = archives[-1][0] if archives else ""
    latest_filename = ""
    latest_diff_label = ""
    if len(archives) >= 2:
        latest_filename = f"{archives[-2][0]}_to_{archives[-1][0]}.html"
        latest_diff_label = (
            f"{period_label(archives[-2][0])} \u2192 {period_label(archives[-1][0])}"
        )
    latest = {
        "period": latest_period,
        "label": period_label(latest_period),
        "diff_label": latest_diff_label,
        "filename": latest_filename,
    }

    reports: list[dict] = []

    if len(archives) < 2:
        print(
            f"Need at least two archives in {tables_dir} to generate a diff report "
            f"(found {len(archives)}). The diff index will be empty.",
            file=sys.stderr,
        )
    else:
        for index in range(1, len(archives)):
            previous_period, previous_zip = archives[index - 1]
            current_period, current_zip = archives[index]
            print(f"Comparing {previous_zip.name} -> {current_zip.name}")

            summary, results = compare_archive_pair(previous_zip, current_zip)
            filename = f"{previous_period}_to_{current_period}.html"

            reports.append(
                {
                    "previous_period": previous_period,
                    "current_period": current_period,
                    "previous_zip": previous_zip.name,
                    "current_zip": current_zip.name,
                    "filename": filename,
                    "files_changed": summary["files_changed"],
                    "added": summary["added"],
                    "removed": summary["removed"],
                    "modified": summary["modified"],
                }
            )

            previous_link = None
            if index > 1:
                previous_link = f"{archives[index - 2][0]}_to_{previous_period}.html"
            next_link = None
            if index < len(archives) - 1:
                next_link = f"{current_period}_to_{archives[index + 1][0]}.html"

            report_html = render_report(
                summary,
                results,
                previous_period,
                current_period,
                previous_link,
                next_link,
                generated_at,
                home_from_diffs,
                latest,
            )
            (output_dir / filename).write_text(report_html, encoding="utf-8", newline="\n")
            print(
                f"  files changed: {summary['files_changed']}, "
                f"added: {summary['added']}, removed: {summary['removed']}, "
                f"modified: {summary['modified']}"
            )

    index_html = render_index(reports, generated_at, home_from_diffs, latest)
    (output_dir / "index.html").write_text(index_html, encoding="utf-8", newline="\n")

    home_html = render_home(
        archives_info,
        reports,
        generated_at,
        tables_link,
        diffs_link,
        notes_link,
        notes_count,
        latest,
    )
    index_file.write_text(home_html, encoding="utf-8", newline="\n")

    print(f"Wrote {len(reports)} report(s) to {output_dir} and landing page to {index_file}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
