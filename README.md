# Email Permission Matching Project

This project provides a repeatable workflow for matching employee records against a set
of Outlook message files and classifying the type of permission change
requested.  It was originally developed to process a roster provided
in `名单.xlsx` and a zip archive of converted `.eml` files (`convert-outlook-msg-file-primary.zip`).

## Overview

The goal of the workflow is to find, for each employee in the roster,
the most recent e‑mail that references either their name or their ID
number.  Based on simple keyword counting, the request is classified
as either “新增” (add), “删除” (remove) or “修改” (modify).  A default
classification of “新增” is applied when no keywords are present or
when counts are tied.  Results are recorded in a new workbook with
two sheets:

* **结果** – records with a matched e‑mail, scenario classification,
  evidence snippet and source (subject, body or attachment).
* **需你决策** – employees for whom no relevant e‑mail could be found.

## Running the Script

Install the Python dependencies and ensure that the Poppler `pdftotext`
utility is available on your system.  Then run the script via:

```sh
python main.py --excel 名单.xlsx --zip convert-outlook-msg-file-primary.zip --output 匹配结果.xlsx
```

### Parameters

* `--excel` – path to the roster spreadsheet.  Must contain a
  column named `员工名字工号` with records formatted as
  `Name, ID` or `ID, Name`.
* `--zip` – path to a zip file containing the `.eml` messages.
* `--output` – path where the resulting workbook should be written.

## Heuristics

The matching and classification logic is intentionally conservative to
avoid false positives:

1. **Eligible Attachments** – Only `.txt`, `.html`, `.csv`, `.docx`,
   `.xlsx` and `.pdf` attachments smaller than 5 MB are parsed.
2. **Manager Filtering** – Occurrences of a name/ID in message bodies
   are ignored if they appear near the word “manager” or within
   quoted header lines (e.g. “From:”, “To:”, “Cc:”); this prevents
   incorrectly matching a manager’s details when they are copied on
   an e‑mail.
3. **Scenario Detection** – Counts of `新增`, `删除` and `修改` are
   computed on the combined text of subject, body and attachments.
   Instances of “修改密码” are removed before counting.  If counts are
   equal or all zero, “新增” is assigned.

## Customisation

The matching heuristics can be tuned by editing `main.py`.  For
example, to adjust how manager context is detected, modify the
`body_has_valid_match` function.  Additional attachment types can
be enabled by adding extensions to `ALLOWED_EXT`.

## License

This project is provided as-is under the MIT License.  See
`LICENSE` for details.