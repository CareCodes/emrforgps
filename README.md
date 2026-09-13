# EmrForGps — Electronic Medical Assistant 3

> **Archived, unmaintained historical project.** This repository is kept for provenance only.
> The actively developed successor is **[hmislk/hmis](https://github.com/hmislk/hmis)** — an
> open source Java EE Hospital Information Management System.

**Electronic Medical Assistant (EMA) version 3** was an Electronic Medical Record system for
General Practitioners, written in **Visual Basic 6** with a **Microsoft Access** backend. It
covered consultations and clinical notes, diagnosis and prescription entry, investigation
ordering, pharmacy stock and batch/expiry tracking, billing, deposits, day-end cash
reconciliation and clinical reporting.

Originally authored by **Dr. M H B Ariyaratne**. The source was previously hosted on Bitbucket
(`bariyaratne/emrforgps`) and has been migrated here so the history of the project is not lost.

## Status

Archived and read-only. It will not build or run on modern toolchains without a Visual Basic 6
IDE/runtime and the 32-bit Microsoft Access (Jet/ACE) OLE DB provider, both long out of support.
No issues, pull requests or support requests are accepted.

## Layout

| Path | Contents |
| --- | --- |
| `EmrForGps.vbp` | Main application VB6 project |
| `Administrator.vbp` | Administrator application VB6 project |
| `*.frm` / `*.frx` | Forms and their binary resources |
| `*.bas` | Modules — find, printing, reporting, security, photo, globals |
| `*.cls` | Class modules |
| `DataEnvironment1.*` | ADO data environment |
| `EmrForGps.PDM` | PowerDesigner data model |
| `EMA3Help.chm` | Compiled user help file |
| `*.txt` | Original development notes and release documentation |

The application database (`EMA3.mdb`) is **deliberately excluded** from this repository: it may
contain real patient data and must not be published.

## Redactions

This archive is not a byte-for-byte copy of the original source. The following were removed
before publication, and the history was rewritten so they appear in no commit:

- The hard-coded Microsoft Jet database password, which appeared in `modStart.bas`,
  `frmSelectDatabase.frm` and `DataEnvironment1.dsx`, and doubled as an `admin` credential for a
  web-hosted copy of the database. Replaced with `REDACTED` / `XXXXXXX`.
- `Mr. Dharmapala.txt`, which contained only an individual's name and mobile number.
- Two third-party staff mobile numbers in `EMA 3 Fintec.txt`.

The application database (`EMA3.mdb`) was never committed: it contains real patient data.
Consequently this source will not connect to a database as-is.

## Licence

Released under the **MIT License** — see [LICENSE](LICENSE).

MIT was chosen as the most permissive licence that already existed when this software was
written; the GNU Affero GPL v3 did not appear until November 2007.

The original demonstration notice that shipped with the software ("all copyrights reserved,
you can't copy and distribute it") has been removed: it is superseded by the MIT License
granted here by the copyright holder.
