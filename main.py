#!/usr/bin/env python3
"""
main.py
========

This script processes a roster of employees and a collection of e‑mail messages
to determine whether there is an associated approval e‑mail for each
employee.  It applies a series of heuristics to classify the request as
``新增`` (add), ``删除`` (remove) or ``修改`` (modify).  Results are written
to an Excel workbook with two worksheets: one for automatically
resolved entries and one for those requiring manual review.

Usage
-----
    python main.py --excel 名单.xlsx --zip convert-outlook-msg-file-primary.zip --output output.xlsx

Parameters
----------
``--excel``: Path to the roster Excel file.  The workbook must contain
  a worksheet with the column ``员工名字工号`` as a comma–separated
  combination of employee name and ID.  Additional columns such as
  ``权限变更场景 (新增，删除，修改)``, ``操作是否正确`` and
  ``新增/删除场景，是否收到申请邮件`` will be updated where appropriate.

``--zip``: Path to a zip archive containing ``*.msg.eml`` files.  The
  archive is extracted to a temporary directory for processing.

``--output``: Path to the resulting Excel workbook.  Two sheets will be
  created: ``结果`` (for resolved entries) and ``需你决策`` (for
  unresolved entries).

Notes
-----
This script avoids external Python dependencies beyond the standard
library and ``pandas``/``openpyxl``.  PDF attachments are converted
using the ``pdftotext`` command-line tool (part of the poppler
utilities), so ``pdftotext`` must be available in your execution
environment.  Text extraction from ``.docx`` files is achieved by
reading the internal ``word/document.xml`` file.  Only attachments of
types ``.txt``, ``.html``, ``.csv``, ``.docx``, ``.xlsx`` and
``.pdf`` are considered, and files larger than 5 MB are skipped.

Heuristics
----------
- A match is considered if either the employee name or ID appears in
  the subject, body, or eligible attachments.  Matches found in
  quoted header sections (``From:``, ``To:``, ``Cc:``, etc.) or within
  200 characters of the word ``manager`` are ignored to prevent
  inadvertently matching a supervisor rather than the employee.
- The scenario (``新增``/``删除``/``修改``) is determined by counting
  occurrences of those keywords in the combined text of subject,
  body and attachments.  Occurrences of ``修改密码`` (change password)
  are ignored.  If no keyword is present or counts are tied,
  ``新增`` is used as a default.
- For each employee there may be multiple matching e‑mails; the
  most recent e‑mail (by Date header) is used.
"""

import argparse
import csv
import io
import os
import re
import shutil
import sys
import tempfile
import zipfile
from collections import defaultdict
from datetime import datetime
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime

import pandas as pd
from bs4 import BeautifulSoup

# -----------------------------------------------------------------------------
# Attachment extraction utilities
# -----------------------------------------------------------------------------
ALLOWED_EXT = {".txt", ".html", ".csv", ".docx", ".xlsx", ".pdf"}
MAX_ATTACHMENT_SIZE = 5 * 1024 * 1024  # 5 MB


def extract_attachment_text(filename: str, data: bytes) -> str:
    """Return textual content from an attachment.

    Only attachments with extensions in ALLOWED_EXT and smaller than
    MAX_ATTACHMENT_SIZE are processed.  For ``.pdf`` files the
    ``pdftotext`` utility is invoked.
    """
    low = filename.lower()
    ext = None
    for e in ALLOWED_EXT:
        if low.endswith(e):
            ext = e
            break
    if ext is None or len(data) > MAX_ATTACHMENT_SIZE:
        return ""

    # Text based attachments
    if ext in {".txt", ".html", ".csv"}:
        for enc in ["utf-8", "gbk", "latin1"]:
            try:
                return data.decode(enc)
            except Exception:
                pass
        return ""

    # Word document (.docx)
    if ext == ".docx":
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
            # remove XML tags
            return re.sub("<[^>]+>", "", xml)
        except Exception:
            return ""

    # Excel spreadsheet (.xlsx)
    if ext == ".xlsx":
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(data)
            tmp_path = tmp.name
        try:
            df_dict = pd.read_excel(tmp_path, sheet_name=None, dtype=str)
            cells: list[str] = []
            for df in df_dict.values():
                cells.extend(df.astype(str).values.flatten().tolist())
            return " ".join(cells)
        except Exception:
            return ""
        finally:
            os.remove(tmp_path)

    # PDF document (.pdf)
    if ext == ".pdf":
        # Write to temp file then convert using pdftotext
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(data)
            pdf_path = tmp.name
        txt_path = pdf_path.replace(".pdf", ".txt")
        try:
            import subprocess

            subprocess.run(["pdftotext", "-layout", pdf_path, txt_path], check=True,
                           stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            with open(txt_path, "r", encoding="utf-8", errors="ignore") as f:
                return f.read()
        except Exception:
            return ""
        finally:
            os.remove(pdf_path)
            if os.path.exists(txt_path):
                os.remove(txt_path)

    return ""


# -----------------------------------------------------------------------------
# Email parsing
# -----------------------------------------------------------------------------

def parse_eml_file(path: str) -> dict:
    """Parse an .eml file and return a dictionary with extracted fields."""
    try:
        with open(path, "rb") as f:
            msg = BytesParser(policy=policy.default).parse(f)
    except Exception:
        return {}

    subject = msg["subject"] or ""
    date_str = msg["date"]
    msg_date: datetime | None = None
    if date_str:
        try:
            dt = parsedate_to_datetime(date_str)
            # convert to naive UTC for comparison
            if dt.tzinfo is not None:
                msg_date = dt.astimezone(datetime.utcfromtimestamp(0).tzinfo).replace(tzinfo=None)
            else:
                msg_date = dt
        except Exception:
            msg_date = None

    # Extract body (plain text preferred; HTML fallback stripped of tags)
    plain_parts: list[str] = []
    html_parts: list[str] = []
    for part in msg.walk():
        if part.get_content_disposition() == "attachment":
            continue
        ctype = part.get_content_type()
        try:
            if ctype == "text/plain":
                plain_parts.append(part.get_content())
            elif ctype == "text/html":
                html_parts.append(part.get_content())
        except Exception:
            continue
    if plain_parts:
        body = "\n".join(plain_parts)
    elif html_parts:
        texts: list[str] = []
        for html in html_parts:
            soup = BeautifulSoup(html, "html.parser")
            texts.append(soup.get_text(separator="\n"))
        body = "\n".join(texts)
    else:
        body = ""

    # Extract eligible attachments' text
    attachments: list[dict] = []
    for att in msg.iter_attachments():
        fname = att.get_filename()
        if not fname:
            continue
        low = fname.lower()
        if not any(low.endswith(ext) for ext in ALLOWED_EXT):
            continue
        data = att.get_payload(decode=True) or b""
        if len(data) > MAX_ATTACHMENT_SIZE:
            continue
        text = extract_attachment_text(low, data)
        attachments.append({"filename": fname, "text": text})

    return {
        "file": os.path.basename(path),
        "subject": subject,
        "body": body,
        "attachments": attachments,
        "date": msg_date,
    }


def parse_emails_from_zip(zip_path: str) -> list[dict]:
    """Extract and parse all .eml files from a zip archive."""
    temp_dir = tempfile.mkdtemp()
    emails: list[dict] = []
    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(temp_dir)
        for name in os.listdir(temp_dir):
            full_path = os.path.join(temp_dir, name)
            if name.lower().endswith(".eml") or name.lower().endswith(".msg.eml"):
                parsed = parse_eml_file(full_path)
                if parsed:
                    emails.append(parsed)
    finally:
        shutil.rmtree(temp_dir)
    return emails


# -----------------------------------------------------------------------------
# Matching and classification logic
# -----------------------------------------------------------------------------

def parse_name_and_id(value: str) -> tuple[str, str]:
    """Split the roster field "员工名字工号" into name and ID.

    The field is expected to be "Name, ID" or "ID, Name".  If the
    first part contains digits it is treated as the ID.  Otherwise the
    second part is assumed to be the ID.  If no comma is present the
    entire field is treated as the name and the ID is left empty.
    """
    s = (value or "").strip()
    if not s:
        return "", ""
    parts = [p.strip() for p in s.split(",")]
    if len(parts) == 2:
        p0, p1 = parts
        if re.search(r"\d", p0):
            return p1, p0
        else:
            return p0, p1
    return s, ""


def body_has_valid_match(body_lower: str, keyword_lower: str) -> bool:
    """Return True if the keyword appears in the body outside of header/manager context.

    - Lines beginning with ``From:``, ``To:``, ``Cc:``, ``Sent:`` or
      ``Subject:`` or containing an ``@`` sign are considered header lines
      and are ignored.
    - A match is rejected if the preceding 200 characters contain the
      word ``manager``.  This reduces false matches where the name
      appears in a manager listing.
    """
    lines = body_lower.split("\n")
    # Check each occurrence line by line
    for line_no, line in enumerate(lines):
        if keyword_lower in line:
            trimmed = line.strip().lower()
            # Skip header/metadata lines
            if trimmed.startswith(("from:", "to:", "cc:", "sent:", "subject:")):
                continue
            if "@" in line:
                continue
            # Confirm there is no manager context immediately before
            # compute absolute position for manager check
            full_text = body_lower
            idx = full_text.find(keyword_lower)
            if idx != -1:
                prefix = full_text[max(0, idx - 200): idx]
                if "manager" in prefix:
                    continue
            return True
    return False


def determine_scenario(text_lower: str) -> str:
    """Classify the scenario by counting keywords.

    Counts the occurrences of ``新增``, ``删除`` and ``修改`` in the
    provided text (with occurrences of ``修改密码`` removed).  If no
    keyword is present or there is a tie for highest count, defaults to
    ``新增``.
    """
    text_lower = text_lower.replace("修改密码", "")
    keywords = ["修改", "删除", "新增"]
    counts = {kw: text_lower.count(kw) for kw in keywords}
    max_val = max(counts.values()) if counts else 0
    if max_val == 0:
        return "新增"
    top = [k for k, v in counts.items() if v == max_val]
    if len(top) == 1:
        return top[0]
    return "新增"


def process_roster(df: pd.DataFrame, emails: list[dict]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Process the roster DataFrame and return resolved and unresolved entries.

    :param df: DataFrame loaded from the roster Excel file
    :param emails: List of dictionaries representing parsed e‑mails
    :return: (resolved_df, unresolved_df)
    """
    resolved = []
    unresolved = []

    # Precompute lower‑case content for each e‑mail
    for email in emails:
        email["subject_lower"] = (email["subject"] or "").lower()
        email["body_lower"] = (email["body"] or "").lower()
        email["attachments_lower"] = [att["text"].lower() if att["text"] else "" for att in email["attachments"]]
        combined = "\n".join([email["subject_lower"], email["body_lower"]] + email["attachments_lower"])
        email["combined_lower"] = combined

    for _, row in df.iterrows():
        name_raw = row.get("员工名字工号", "")
        name, eid = parse_name_and_id(name_raw)
        name_low = name.lower()
        id_low = eid.lower()
        matches: list[tuple[dict, str]] = []  # list of (email, source)

        for email in emails:
            source = None
            # Check subject
            if name_low and name_low in email["subject_lower"] or id_low and id_low in email["subject_lower"]:
                source = "主题"
            # Check body
            if not source and (name_low or id_low):
                body_low = email["body_lower"]
                found = False
                if name_low and name_low in body_low and body_has_valid_match(body_low, name_low):
                    found = True
                elif id_low and id_low in body_low and body_has_valid_match(body_low, id_low):
                    found = True
                if found:
                    source = "正文"
            # Check attachments
            if not source and (name_low or id_low):
                for att_text in email["attachments_lower"]:
                    if name_low and name_low in att_text or id_low and id_low in att_text:
                        source = "附件"
                        break
            if source:
                matches.append((email, source))

        if matches:
            # Pick the most recent e‑mail (latest date)
            matches.sort(key=lambda x: x[0]["date"] or datetime.min, reverse=True)
            selected_email, source = matches[0]
            scenario = determine_scenario(selected_email["combined_lower"])
            # Determine snippet for evidence
            snippet_source_text = ""
            if source == "正文":
                snippet_source_text = selected_email["body"]
            elif source == "附件":
                # find the first attachment that contains the keyword
                for att in selected_email["attachments"]:
                    text_low = att["text"].lower() if att["text"] else ""
                    if name_low and name_low in text_low or id_low and id_low in text_low:
                        snippet_source_text = att["text"]
                        break
                if not snippet_source_text:
                    snippet_source_text = selected_email["body"]
            else:  # subject match
                snippet_source_text = selected_email["body"] if selected_email["body"] else selected_email["subject"]
            snippet = snippet_source_text[:500] if snippet_source_text else ""

            result_row = row.copy()
            orig_scen = str(row.get("权限变更场景 (新增，删除，修改)") or "")
            if not orig_scen or orig_scen.lower() in {"nan", ""}:
                result_row["权限变更场景 (新增，删除，修改)"] = scenario
                result_row["操作是否正确"] = "是"
            # If scenario is 新增 or 删除, mark receipt
            if scenario in {"新增", "删除"}:
                result_row["新增/删除场景，是否收到申请邮件"] = "是"
            result_row["最终判定来源"] = source
            result_row["匹配说明"] = snippet
            resolved.append(result_row)
        else:
            # No match; needs manual decision
            manual_row = row.copy()
            manual_row["最终判定来源"] = ""
            manual_row["匹配说明"] = "未匹配到邮件"
            unresolved.append(manual_row)

    return pd.DataFrame(resolved), pd.DataFrame(unresolved)


# -----------------------------------------------------------------------------
# CLI entry point
# -----------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Process roster and e‑mail archive")
    parser.add_argument("--excel", required=True, help="Path to the roster Excel file")
    parser.add_argument("--zip", required=True, help="Path to the zip archive of e‑mails")
    parser.add_argument("--output", required=True, help="Path of the output Excel workbook")
    args = parser.parse_args()

    # Read roster
    try:
        df = pd.read_excel(args.excel, dtype=str)
    except Exception as exc:
        print(f"Error reading roster: {exc}", file=sys.stderr)
        sys.exit(1)
    # Validate required column
    if "员工名字工号" not in df.columns:
        print("Roster must contain a column named '员工名字工号'", file=sys.stderr)
        sys.exit(1)

    # Parse emails
    print("Extracting and parsing e‑mails…")
    emails = parse_emails_from_zip(args.zip)
    print(f"Parsed {len(emails)} e‑mails")

    # Process roster
    resolved_df, manual_df = process_roster(df, emails)
    print(f"Resolved entries: {len(resolved_df)}, Unresolved: {len(manual_df)}")

    # Write results
    with pd.ExcelWriter(args.output) as writer:
        resolved_df.to_excel(writer, sheet_name="结果", index=False)
        manual_df.to_excel(writer, sheet_name="需你决策", index=False)
    print(f"Results saved to {args.output}")


if __name__ == "__main__":
    main()