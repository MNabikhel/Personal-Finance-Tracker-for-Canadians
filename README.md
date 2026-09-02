# Canadian Finance Tracker

A personal finance tracker for Canadian households, built as a single
macro-enabled Excel workbook. Download the statements your bank and credit
cards already give you — CSV or PDF — press one button, and the workbook sorts
every transaction into a category and shows you where the money went — for one
person, or for a couple.

**Download the workbook:** [`dist/Canadian-Finance-Tracker.xlsm`](dist/Canadian-Finance-Tracker.xlsm)

Nothing leaves your computer. There is no add-in to install, no account to
create and no connection to your bank — the workbook only ever reads the
statement files you point it at.

---

## Contents

- [What it does](#what-it-does)
- [Getting started](#getting-started)
- [Importing statements](#importing-statements)
- [Categories and rules](#categories-and-rules)
- [Couples mode](#couples-mode)
- [The sheets](#the-sheets)
- [The Canadian bits](#the-canadian-bits)
- [Building it yourself](#building-it-yourself)
- [Tests](#tests)
- [How the repository fits together](#how-the-repository-fits-together)
- [Limits and caveats](#limits-and-caveats)

---

## What it does

**Imports your statements.** Seventeen CSV layouts are pre-configured, covering
RBC, TD, CIBC, Scotiabank, BMO, National Bank, Desjardins, Tangerine, Simplii,
EQ Bank, PC Financial, KOHO, Wealthsimple Cash and Amex Canada, plus two
generic layouts and a blank row to fill in for anything else. The importer
recognises a file from its own header line, shows you the first rows it parsed
and asks you to confirm before writing anything. PDF statements are read too,
through Excel's own PDF reader, and go through the same preview and the same
duplicate check — see [Importing statements](#importing-statements) for what
that needs and where it stops.

**Sorts it out for you.** 194 seeded rules map merchant names onto 92
categories in 14 groups. Anything the rules cannot place is left as
`Uncategorized` and highlighted; select a row, press *Teach a rule*, and every
future transaction like it is categorised automatically.

**Never double-counts.** Re-importing a statement that overlaps one you already
loaded adds only the new rows, while two genuinely identical purchases on the
same day are still kept as two. Transfers between your own accounts and credit
card payments are matched up by amount and date and excluded from both income
and expenses.

**Tells you where you stand.** A dashboard with money in / money out / saved,
net cash flow, savings rate, essential versus discretionary spending, your
biggest merchants and three charts; a twelve-month report by group and by
category; and a budget sheet that compares each category against the monthly
budget you set.

**Works for two people, if you want it to.** In couple mode every transaction
gets an owner — one of you, or Joint. Joint costs are divided by your household
split, the Household sheet works out who owes whom, and a selector on the
dashboard switches every number in the workbook between the whole household and
one person's share. Answer "Just me" in the setup wizard and all of it
disappears.

**Speaks Canadian.** Hydro, Presto/Compass/OPUS, LCBO/SAQ, property tax,
daycare, OSAP, CRA payments and refunds, the Canada Child Benefit, the GST/HST
credit and the Canada Carbon Rebate all have categories and rules. There is a
tax-summary sheet keyed to CRA line numbers and a registered-plans sheet that
tracks RRSP/TFSA/FHSA/RESP contributions per person against your room.

The file ships with six months of realistic sample data (324 transactions
across four accounts) so that everything is populated the first time you open
it. The setup wizard offers to delete it when you are ready for your own.

---

## Getting started

1. **Download** [`dist/Canadian-Finance-Tracker.xlsm`](dist/Canadian-Finance-Tracker.xlsm)
   and save it somewhere you keep backups. On Windows, right-click the file,
   choose *Properties* and tick **Unblock** — files from the internet open in
   Protected View with macros disabled otherwise.
2. **Open it in Excel and choose *Enable Content*.** The import, the
   categorising and the buttons are all macros; without them you get a
   spreadsheet full of sample data and nothing else.
3. **Look around.** The dashboard opens on the last complete month of the
   sample data, and every button works against it.
4. **Press *Setup wizard*.** It asks whether the workbook is for one person or
   two, your names, how you split shared costs, and your province — then offers
   to clear the sample transactions.
5. **List your accounts** on the *Accounts* sheet, one row per bank or card
   account. Put a snippet of the file name your bank produces in *File Name
   Contains* and imports will find the right account by themselves.
6. **Download a statement from each bank — CSV or PDF — and press *Import
   statements*.**

### If Excel complains when you open it

- **A red bar: "SECURITY RISK — Microsoft has blocked macros from running
  because the source of this file is untrusted."** Windows marks anything
  downloaded from the internet, and Excel refuses macros in marked files. Close
  the workbook, right-click the file, choose *Properties*, tick **Unblock**,
  *OK*, and open it again. Alternatively keep it in a folder you have added as
  a Trusted Location (*File › Options › Trust Center › Trust Center Settings ›
  Trusted Locations*).
- **A yellow bar: "Protected View"** — press *Enable Editing*, then *Enable
  Content* on the bar that follows.
- **"Excel cannot open the file because the file format or file extension is
  not valid."** The download saved GitHub's web page rather than the workbook.
  Use the *Download raw file* button on the file's GitHub page (or clone the
  repository); the real file is about 540 KB and opens as a zip archive.
- **"We found a problem with some content … Do you want us to try to
  recover?"** means a part of the file is malformed. The build checks the
  parts Excel is strict about — table headers, defined names, the VBA project
  binary — and this should not happen with a current build; if you see it,
  please open an issue with the Excel version and the text of the repair
  report, and in the meantime rebuild from source (below).

### Buttons and shortcuts

The dashboard carries *Import statements*, *Apply rules*, *Find transfers*,
*Needs a category*, *Refresh*, *Setup wizard*, *Couple mode on/off* and *Help*.
The Transactions sheet carries *Import statements*, *Undo an import*, *Teach a
rule*, *Set owner*, *Show all rows*, *Apply rules to all*, *Rebuild formulas*,
*Start fresh* and *Back to dashboard*.

| Shortcut | Does |
| --- | --- |
| `Ctrl+Shift+I` | Import statements |
| `Ctrl+Shift+R` | Apply rules to uncategorised rows |
| `Ctrl+Shift+T` | Teach a rule from the selected row |
| `Ctrl+Shift+U` | Show rows that still need a category |

The buttons are drawn by a macro when the file opens rather than saved into it,
so if they are ever missing, close and reopen the workbook with macros enabled.

---

## Importing statements

*Import statements* takes CSV exports and PDF statements, any mix of them at
once. CSV is the surer of the two — it is data, and every Canadian bank offers
it from the transaction list — so it is described first; PDFs are
[below](#pdf-statements).

### CSV exports

These layouts are pre-configured on the *Bank Formats* sheet:

| Institution | Profile |
| --- | --- |
| RBC Royal Bank | RBC Chequing/Savings/Card |
| TD Canada Trust | TD (no header) |
| CIBC | CIBC (no header) |
| Scotiabank | Scotiabank (no header) |
| BMO | BMO |
| National Bank of Canada | National Bank |
| Desjardins | Desjardins |
| Tangerine | Tangerine |
| Simplii Financial | Simplii (no header) |
| EQ Bank | EQ Bank |
| PC Financial | PC Financial Mastercard |
| KOHO | KOHO |
| Wealthsimple | Wealthsimple Cash |
| American Express | Amex Canada |
| Any | Generic: date, description, amount |
| Any | Generic: date, description, debit, credit |
| Your bank | Custom (edit me) |

A profile is thirteen cells on a worksheet — how many rows to skip, the
delimiter, which column holds the date and in what format, which columns make
up the description, whether amounts are signed or split into debit and credit
columns, and a fragment of the header line to recognise the file by. If your
bank changes its export or is not listed, fix the column numbers on that sheet
and import again. No code changes, no rebuild.

Three things worth knowing:

- **Sign convention.** Money leaving an account is stored as a negative number
  and money arriving as a positive one, whichever way your bank chose to write
  it. Amex, for instance, reports purchases as positive; the Amex profile flips
  them.
- **Header rows and bank notices are skipped, not fatal.** A BMO export opens
  with three lines of prose; anything that does not parse as a transaction is
  counted as an unreadable row and the rest of the file still imports.
- **Every batch is logged.** The *Import Log* sheet records the file, the
  profile used, the account, and how many rows were read, added, skipped as
  duplicates and skipped as unreadable.
- **Every batch can be taken back.** *Undo an import* on the Transactions
  sheet lists the recent batches, deletes everything the one you pick added,
  and marks its log row *Undone*. Wrong file, wrong account, or a PDF whose
  signs came out the wrong way round: one press, not a hunt through the
  ledger. Transaction numbers are never reused afterwards.

### PDF statements

A PDF is a printed page, not data: the file stores where each piece of text
sits on the page, and a reader has to work out which pieces form a row and
which a column. Excel has exactly one such reader, Power Query's *From PDF*
connector, and the workbook drives it from VBA: the PDF is loaded through a
temporary query onto a scratch sheet, its lines are read, and the query and
the sheet are removed again. Where that reader is missing, Word is asked to
convert the file instead (any Word since 2013 opens PDFs). Where neither is
available, the import says so and points you at the CSV.

What that means in practice:

| | |
| --- | --- |
| **Works in** | Excel for Windows under Microsoft 365 (Power Query's PDF reader), or any Excel on a machine with Word 2013 or later. |
| **Does not work in** | Excel for Mac, Excel on the web, and perpetual Excel 2016/2019 without Word. Import the CSV there. |
| **Scanned statements** | Have no text in them and cannot be read. The import tells you so. |

Once the text is out, a line that starts with a date and ends with an amount
is a transaction, and everything else on the page — addresses, totals,
notices, column headings — is not. Card statements print a transaction date
and a posting date; the first is used. A plain amount on a card is a charge,
a minus sign or a `CR` a credit. Bank account statements carry a running
balance, and the balance moving down or up is what tells a withdrawal from a
deposit; a line printed without a date (RBC prints the date once for a day's
transactions) takes the date above it, provided the balance vouches for it.
Reference numbers ahead of the date (Scotiabank) are stepped over, French
statements with `03 mars` and `1 234,56` are read, and the year — which card
statements do not print beside each line — comes from the latest full date on
the page, with December lines on a January statement placed in the year
before. The importer decides whether it is looking at a card or a bank
statement from the words on the page or, failing those, from whether the
balances add up, and offers the other reading if you say no to the preview.

The preview matters more for PDFs than for CSVs. Look at the signs before you
say yes; if the bank prints its statement in some shape this does not read,
the CSV always works. The first time Power Query reads a file it may ask about
privacy levels; answering *Public* or ignoring them for this file is fine,
since the query only ever reads the one PDF and writes nowhere.

---

## Categories and rules

Rules run in priority order and the first match wins, so specific rules should
have a lower number. Each rule matches on the merchant, the description, the
account or any of them, using *Contains*, *Starts With*, *Ends With*, *Equals*,
*Like* or *Word*, and can be limited to money in, money out, or an amount
range. *Word* matches only at word boundaries, which is what keeps a fuel rule
for `MOBIL` off Freedom Mobile.

Refunds take care of themselves. A credit that no rule claims is tried once
more against the merchant rules as though it were the purchase, so a return at
Loblaws lands in *Groceries* as money in and nets off the spend. Deposits never
get that far: the income and transfer rules go by the description and run
first.

Merchant names are cleaned before matching: payment-processor prefixes, store
numbers, reference numbers and city/province suffixes are stripped, and
Canadian acronyms like TFSA and LCBO survive title-casing intact.

A category you type in by hand is tagged *Manual* and is never overwritten by
*Apply rules to all*.

---

## Couples mode

The sample data is a couple's, so the file arrives in couple mode; the setup
wizard asks which you are and switches the whole workbook to match. After that
you can change your mind at any time with the *Couple mode on/off* button on
the dashboard or the drop-down on the *Settings* sheet. Single mode only hides
what needs two people — the per-person ledger columns, the Household sheet and
the couple block on the dashboard. Nothing is deleted, and going back and forth
is safe.

With it on:

- Every transaction has an **Owner**: one of you, or Joint. Joint transactions
  are divided by the household split you chose, which any category can override
  in *Joint Split A* (daycare 50/50 but a personal hobby 100/0, say).
- **Paid By** is derived from the account the money actually left, so it is
  possible — and normal — to pay for something that is not all yours.
- The **Household** sheet compares what each of you paid from your own accounts
  against your fair share and states the settlement in words ("Sam owes Alex
  $412.60", say) rather than leaving you to interpret a sign. It also shows
  income shares and what an income-proportional split would look like, both for
  the month and year to date.
- The **View** selector on the dashboard switches the entire workbook between
  the household total and either person's share.

---

## The sheets

| Sheet | What it is for |
| --- | --- |
| **Dashboard** | The month at a glance: money in, out and saved, net cash flow, savings rate, essential versus discretionary, spending by group, biggest merchants, three charts, and the buttons. |
| **Transactions** | The ledger. One row per transaction, 24 columns, most of them formulas. This is the only sheet holding your data. |
| **Accounts** | One row per bank or card account: institution, type, owner, and the file-name fragment used to route imports. |
| **Categories** | The 92 categories with their group, type, whether they are essential, tax tag, monthly budget, default owner and joint split. |
| **Rules** | The categorisation rules, in priority order, with a hit counter. |
| **Budget** | Every category against the monthly budget you set, for the report month: budget, actual, difference, % used and a twelve-month average. |
| **Reports** | Twelve months ending on the report month, by group and by category, with totals. |
| **Household** | Couple mode only: settling up, income shares, year to date. |
| **Tax Summary** | Totals per CRA line number for the tax year of the report month, split per person. |
| **Registered Plans** | RRSP/TFSA/FHSA/RESP contributions per person against your room, with the published annual limits. |
| **Bank Formats** | The CSV layouts described above. Editable. PDFs do not use it. |
| **Import Log** | One row per imported file, CSV or PDF, and for a PDF which reader got its text out. |
| **Settings** | Household mode, names, split, province, transfer window, duplicate skipping. |
| **Help** | The same guidance as this README, inside the workbook. |
| **Engine** | Hidden. Drop-down lists and formula templates. Leave it alone. |

---

## The Canadian bits

- **Categories** for what a Canadian household actually pays: electricity /
  hydro, natural gas, Presto/Compass/OPUS transit, LCBO/SAQ, property tax,
  childcare and daycare, OSAP, and CRA payments and refunds.
- **Benefit deposits** — Canada Child Benefit, GST/HST credit, Canada Carbon
  Rebate, OAS, CPP, EI — are recognised as income rather than mystery
  transfers.
- **Tax tags** map categories onto the lines you will actually fill in: medical
  expenses (33099), donations (34900), child care (21400), tuition (Schedule
  11), union and professional dues (21200), student loan interest (31900),
  support payments (22000), RRSP (Schedule 7), FHSA (Schedule 15), business,
  investment and rental income and expenses.
- **Registered plans** track RRSP, TFSA, FHSA and RESP contributions per
  person, with the published annual dollar limits by year.

None of this is tax or financial advice. The CRA figures in the workbook were
checked in September 2026; confirm anything that matters against canada.ca or
your CRA account.

---

## Building it yourself

The `.xlsm` is generated, not hand-edited. Everything about it — the sheets,
formulas, formatting, sample data and the VBA project — comes out of this
repository.

```bash
pip install -r requirements.txt
python3 build.py
```

That writes `dist/Canadian-Finance-Tracker.xlsm` and refreshes the sample CSVs
under `samples/`. The sample transactions are dated relative to today so a
freshly built workbook always opens on a month with data in it.

Builds are reproducible when you pin the date:

```bash
python3 build.py --today 2026-09-01
```

Two runs of the same source with the same `--today` produce byte-identical
files, which is asserted by the test suite. The committed workbook is the
output of exactly that command.

Other options: `--out PATH` to write elsewhere, `--no-samples` to skip the CSVs.

### Why a build script

Excel cannot store a macro-enabled workbook that was never opened in Excel, so
`tools/` builds one from the specifications instead. `tools/cfb.py` writes an
OLE compound file ([MS-CFB]), `tools/ovba.py` implements the VBA storage format
([MS-OVBA]) including its compression and the `dir` stream, `tools/vbaproject.py`
assembles the modules in `vba/` into a `vbaProject.bin`, `tools/workbook.py`
builds the spreadsheet with openpyxl, and `tools/package.py` splices the two
together into a valid `.xlsm`. Editing the VBA is therefore a matter of editing
plain text files in `vba/` and re-running the build — the macros are in version
control and diffable, which they would not be if the workbook were the source
of truth.

---

## Tests

```bash
pip install -r requirements-dev.txt
python3 -m unittest discover -s tests -t .
```

130 tests, about 15 seconds. They fall into six groups:

- **Format conformance** (`test_ovba.py`) — the compression and encryption
  vectors from [MS-OVBA] §3.2 and §2.3.1.15–17, so the writer is checked
  against Microsoft's own reference data; the finished `vbaProject.bin` is
  read back with `oletools` as an independent parser; and the `dir` stream is
  checked for the two things that make Excel drop a project it can otherwise
  read — a class module recorded as a procedural one, and a reference list
  that repeats the host's own libraries.
- **VBA unit tests** (`test_vba.py`) — the pure functions (date and amount
  parsing, CSV splitting, merchant cleanup, rule matching, hashing) executed
  for real. There is no VBA interpreter available outside Excel, so
  `tests/vbahost.py` loads the actual module sources into headless LibreOffice
  Basic and runs them there.
- **End-to-end import** (`test_import.py`) — the shipped import path over the
  four sample bank exports, checking that every row comes back with the right
  date, sign, description, category, owner and duplicate key, that header rows
  are skipped, that re-importing adds nothing, and that a refund follows its
  purchase into the same category while deposits do not.
- **PDF statement text** (`test_pdftext.py`) — the statement parser, in
  LibreOffice Basic, over text in the shapes Canadian statements take: a card
  with posting dates, `CR` credits and the year only in the header; a chequing
  account with a running balance and the date printed once per day; a
  Scotiabank card with reference numbers; a French statement; a January
  statement with December lines. Getting the text out of the PDF is Excel's
  part and cannot be run here, so `modPdf.bas` is covered by static analysis
  only.
- **The built workbook** (`test_workbook.py`) — the package is a valid zip with
  the VBA project declared and macro-enabled content types; the sheets, tables,
  columns and named ranges the macros reference all exist; table headers match
  their table parts and names and validations are stored the way Excel stores
  them (no leading `=`); and the workbook is
  opened in LibreOffice Calc, recalculated, and its dashboard, reports and
  household figures compared against the same totals computed independently in
  Python.
- **Static analysis** (`test_vbasource.py`) — Excel is the only thing that
  compiles this code and it is not available here, so these tests stand in for
  the compiler. Every qualified and unqualified call must resolve to a public
  procedure taking that many arguments, every constant and class member must
  exist, the sheet, table, column and named-range constants must be the ones
  the builder actually writes, and every macro a button or shortcut names must
  be a public parameterless Sub with room on the sheet for the button bar.

LibreOffice is doing real work here: it is a completely separate implementation
of both OOXML and the Basic dialect VBA is derived from, so agreeing with it is
meaningful evidence rather than our writer agreeing with our reader. The tests
need `libreoffice-calc` and `python3-uno` installed; on Debian or Ubuntu:

```bash
sudo apt-get install libreoffice-calc python3-uno
```

---

## How the repository fits together

```
build.py                  Build entry point
tools/
  cfb.py                  OLE compound file writer [MS-CFB]
  ovba.py                 VBA storage format [MS-OVBA]
  vbaproject.py           Assembles vba/ into vbaProject.bin
  workbook.py             Builds the spreadsheet (openpyxl)
  package.py              Splices the two into an .xlsm
  data.py                 Categories, rules, bank formats, CRA figures
  sample.py               Deterministic sample transactions
vba/
  ThisWorkbook.cls        Workbook events
  modImport.bas           Statement import
  modPdf.bas              Getting the text out of a PDF (Power Query, then Word)
  modPdfText.bas          Statement text to transactions
  modParse.bas            CSV and field parsing
  modProfiles.bas         Bank format detection
  modRules.bas            Merchant cleanup and categorisation
  modTransfers.bas        Internal transfer detection
  modAccounts.bas         Account resolution and import logging
  modHousehold.bas        Couple mode
  modLedger.bas           Writing to the transactions table
  modReport.bas           Refresh and navigation
  modSetup.bas            Setup wizard and housekeeping
  modUI.bas               Buttons and shortcuts
  modUtil.bas             Shared helpers
  modConst.bas            Sheet, table and column names
  clsTxn.cls              One transaction being imported
  clsProfile.cls          One bank format
  clsRule.cls             One categorisation rule
samples/                  Four sample bank exports, in each bank's real shape
tests/                    See above
dist/                     The built workbook
```

`vba/modConst.bas` and `tools/workbook.py` have to agree about every sheet,
table, column and named range; `test_vbasource.py` fails the build if they
drift apart.

---

## Limits and caveats

- **Excel, with macros enabled.** The workbook is built for desktop Excel.
  Excel on the web and the mobile apps do not run VBA, and LibreOffice Calc
  will open the file and recalculate it correctly but will not run the macros
  without VBA compatibility mode.
- **The macros have not been run in Excel itself.** They are exercised in
  LibreOffice Basic and checked by static analysis, which catches a great deal,
  but nothing here has been through the real VBA compiler. That goes double
  for the PDF plumbing: Power Query and Word exist only in Office, so the code
  that drives them has been read carefully but never run.
- **PDF import needs Windows.** Excel's PDF reader is in Excel for Windows
  under Microsoft 365; Word can stand in for it. Elsewhere, and for scanned
  statements, use the CSV. There are no bank connections and no Excel-format
  statements.
- **Bank formats change.** The profiles are a best-effort mapping. Check the
  preview on first import and adjust the *Bank Formats* row if a bank has moved
  its columns.
- **20,000 ledger rows.** That is how far the drop-downs and the conditional
  formatting reach. Rows beyond it still calculate, they just lose the
  decoration.
- **Not tax or financial advice.** The Canadian figures were checked in
  September 2026 and are there to help you organise, not to file.

[MS-CFB]: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/
[MS-OVBA]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/
