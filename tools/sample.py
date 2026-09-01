"""Deterministic sample data.

The workbook ships with a few months of realistic Canadian transactions for a
two-person household so that every chart has something to show.  The same
records are also written out as bank-style CSV files under ``samples/`` so the
import flow can be tried out end to end - and because the duplicate keys match,
importing them over the pre-loaded data is correctly detected as duplicates.
"""

from __future__ import annotations

import random
from dataclasses import dataclass
from datetime import date, timedelta
from typing import Dict, List

from . import data

MONTHS_OF_HISTORY = 6

ACCOUNT_ALEX = "RBC Chequing (Alex)"
ACCOUNT_SAM = "Tangerine Chequing (Sam)"
ACCOUNT_JOINT = "BMO Joint Chequing"
ACCOUNT_CARD = "Amex Cobalt (Joint)"

PERSON_A = "Alex"
PERSON_B = "Sam"


@dataclass
class Txn:
    when: date
    account: str
    owner: str
    description: str
    merchant: str
    amount: float
    category: str

    @property
    def month(self) -> str:
        return self.when.strftime("%Y-%m")


def fnv1a(text: str) -> str:
    """Mirror of modUtil.HashText so keys agree with the macros."""
    high, low = 0x811C, 0x9DC5
    for char in text:
        low ^= ord(char) & 0xFFFF
        product_low = (low & 0xFFFF) * 0x193
        product_high = (high & 0xFFFF) * 0x193 + ((low & 0xFFFF) * 0x100)
        low = product_low & 0xFFFF
        high = (product_high + (product_low >> 16)) & 0xFFFF
    return f"{high:04X}{low:04X}"


def condense(text: str) -> str:
    return " ".join(text.split())


def match_key(account: str, when: date, amount: float, description: str) -> str:
    return fnv1a(
        f"{account.upper()}|{when.strftime('%Y-%m-%d')}|{amount:.2f}|"
        f"{condense(description).upper()}"
    )


# (merchant, description template, category, low, high, times per month)
GROCERY = [
    ("Loblaws", "IDP PURCHASE - {ref} LOBLAWS #{store} TORONTO ON", "Groceries", 48, 165, 3),
    ("No Frills", "IDP PURCHASE - {ref} NO FRILLS #{store} TORONTO ON", "Groceries", 35, 120, 2),
    ("Costco Wholesale", "IDP PURCHASE - {ref} COSTCO WHOLESALE #{store} ON", "Groceries", 90, 280, 1),
]

DAILY = [
    ("Tim Hortons", "IDP PURCHASE - {ref} TIM HORTONS #{store} TORONTO ON", "Coffee & Snacks", 3, 14, 6),
    ("Starbucks", "IDP PURCHASE - {ref} STARBUCKS #{store} TORONTO ON", "Coffee & Snacks", 5, 12, 2),
    ("Uber Eats", "AMZ*UBER EATS TORONTO ON", "Restaurants & Takeout", 22, 68, 2),
    ("Swiss Chalet", "IDP PURCHASE - {ref} SWISS CHALET #{store} ON", "Restaurants & Takeout", 28, 74, 1),
    ("Pizza Nova", "IDP PURCHASE - {ref} PIZZA NOVA #{store} ON", "Restaurants & Takeout", 18, 45, 1),
    ("Shoppers Drug Mart", "IDP PURCHASE - {ref} SHOPPERS DRUG MART #{store} ON", "Prescriptions & Pharmacy", 12, 85, 1),
    ("Lcbo", "IDP PURCHASE - {ref} LCBO/RAO #{store} TORONTO ON", "Alcohol", 22, 78, 1),
    ("Presto", "PRESTO FARE/TRANSIT TORONTO ON", "Public Transit", 25, 156, 1),
    ("Petro-Canada", "IDP PURCHASE - {ref} PETRO-CANADA #{store} ON", "Fuel", 45, 95, 2),
    ("Dollarama", "IDP PURCHASE - {ref} DOLLARAMA #{store} ON", "Miscellaneous", 6, 32, 1),
    ("Canadian Tire", "IDP PURCHASE - {ref} CANADIAN TIRE #{store} ON", "Auto Maintenance", 25, 180, 1),
    ("Winners", "IDP PURCHASE - {ref} WINNERS #{store} ON", "Clothing & Shoes", 30, 145, 1),
    ("Indigo", "IDP PURCHASE - {ref} INDIGO #{store} ON", "Hobbies", 15, 70, 1),
    ("Uber", "UBER TRIP HELP.UBER.COM ON", "Taxi & Rideshare", 12, 42, 2),
    ("Home Depot", "IDP PURCHASE - {ref} HOME DEPOT #{store} ON", "Home Improvement", 25, 220, 1),
]

MONTHLY_FIXED = [
    ("Toronto Hydro", "TORONTO HYDRO BILL PAYMENT", "Electricity / Hydro", 78, 142, 8),
    ("Enbridge Gas", "ENBRIDGE GAS PREAUTHORIZED DEBIT", "Natural Gas / Heating", 42, 165, 9),
    ("Rogers", "ROGERS PREAUTHORIZED PAYMENT", "Internet", 89.99, 89.99, 12),
    ("Freedom Mobile", "FREEDOM MOBILE PREAUTHORIZED DEBIT", "Mobile Phone", 45, 45, 14),
    ("Netflix", "NETFLIX.COM MISC PAYMENT", "Subscriptions", 22.99, 22.99, 16),
    ("Spotify", "SPOTIFY P2C5 MISC PAYMENT", "Subscriptions", 16.99, 16.99, 17),
    ("Goodlife Fitness", "GOODLIFE CLUBS PREAUTHORIZED DEBIT", "Fitness & Sports", 54.22, 54.22, 3),
    ("Belairdirect", "BELAIRDIRECT INSURANCE PREAUTHORIZED DEBIT", "Auto Insurance", 168, 168, 5),
    ("Sunlife", "SUN LIFE ASSURANCE MISC PAYMENT", "Life & Disability Insurance", 74.5, 74.5, 6),
]


def _amount(rng: random.Random, low: float, high: float) -> float:
    if low == high:
        return -round(low, 2)
    return -round(rng.uniform(low, high), 2)


def _ref(rng: random.Random) -> Dict[str, str]:
    return {
        "ref": str(rng.randrange(1000, 9999)),
        "store": str(rng.randrange(100, 4999)),
    }


def build(today: date | None = None, seed: int = 20260901) -> List[Txn]:
    rng = random.Random(seed)
    today = today or date.today()
    first_month = (today.replace(day=1) - timedelta(days=1)).replace(day=1)
    months = []
    cursor = first_month
    for _ in range(MONTHS_OF_HISTORY):
        months.append(cursor)
        year = cursor.year - 1 if cursor.month == 1 else cursor.year
        month = 12 if cursor.month == 1 else cursor.month - 1
        cursor = date(year, month, 1)
    months.reverse()

    out: List[Txn] = []

    def add(when: date, account: str, owner: str, description: str, merchant: str,
            amount: float, category: str) -> None:
        # A shared expense stays shared no matter whose card paid for it - the
        # same rule the workbook applies when it categorises an import.
        owner = data.default_owner(category) or owner
        out.append(Txn(when, account, owner, condense(description), merchant,
                       round(amount, 2), category))

    for month_start in months:
        days_in_month = _days_in_month(month_start)

        # Pay cheques, twice a month each.
        for day, who, account, gross in (
            (15, PERSON_A, ACCOUNT_ALEX, 2465.18),
            (30, PERSON_A, ACCOUNT_ALEX, 2465.18),
            (15, PERSON_B, ACCOUNT_SAM, 1980.44),
            (30, PERSON_B, ACCOUNT_SAM, 1980.44),
        ):
            when = month_start.replace(day=min(day, days_in_month))
            employer = "NORTHWIND LOGISTICS" if who == PERSON_A else "MAPLE HEALTH GROUP"
            add(when, account, who, f"PAYROLL DEPOSIT {employer}", employer.title(),
                gross + rng.randrange(-40, 41), "Employment Income")

        # Canada Child Benefit lands on the 20th.
        when = month_start.replace(day=min(20, days_in_month))
        add(when, ACCOUNT_JOINT, "Joint", "CANADA FED / FED CCB DEPOSIT", "Canada Fed",
            619.75, "Canada Child Benefit")

        # Rent and childcare come out of the joint account.
        add(month_start.replace(day=1), ACCOUNT_JOINT, "Joint",
            "RENT PAYMENT 84 BROADVIEW AVE", "Rent Payment", -2350, "Rent")
        add(month_start.replace(day=min(5, days_in_month)), ACCOUNT_JOINT, "Joint",
            "LITTLE MAPLE DAYCARE PREAUTHORIZED DEBIT", "Little Maple Daycare",
            -880, "Childcare & Daycare")

        # Regular bills, spread through the month.
        for merchant, template, category, low, high, day in MONTHLY_FIXED:
            when = month_start.replace(day=min(day, days_in_month))
            add(when, ACCOUNT_JOINT, "Joint", template, merchant,
                _amount(rng, low, high), category)

        # Savings and investing.
        add(month_start.replace(day=min(16, days_in_month)), ACCOUNT_ALEX, PERSON_A,
            "TRANSFER TO RRSP WEALTHSIMPLE", "Wealthsimple Rrsp", -450,
            "RRSP Contribution")
        add(month_start.replace(day=min(16, days_in_month)), ACCOUNT_SAM, PERSON_B,
            "TFSA CONTRIBUTION TANGERINE INVESTMENT", "Tangerine Tfsa", -350,
            "TFSA Contribution")
        add(month_start.replace(day=min(18, days_in_month)), ACCOUNT_JOINT, "Joint",
            "RESP CONTRIBUTION CST SAVINGS", "Cst Savings", -208.33,
            "RESP Contribution")

        # Everyday spending, split between the personal cards and the joint card.
        for group in (GROCERY, DAILY):
            for merchant, template, category, low, high, times in group:
                for _ in range(times):
                    when = month_start.replace(day=rng.randrange(1, days_in_month + 1))
                    account, owner = _pick_account(rng, category)
                    add(when, account, owner, template.format(**_ref(rng)), merchant,
                        _amount(rng, low, high), category)

        # Card payment from the joint account, plus a couple of bank fees.
        card_total = sum(
            txn.amount for txn in out
            if txn.account == ACCOUNT_CARD and txn.month == month_start.strftime("%Y-%m")
        )
        if card_total:
            when = month_start.replace(day=min(26, days_in_month))
            add(when, ACCOUNT_JOINT, "Joint", "AMEX CREDIT CARD/LOC PAY",
                "Amex Credit Card", card_total, "Credit Card Payment")
            add(when, ACCOUNT_CARD, "Joint", "PAYMENT - THANK YOU / PAIEMENT - MERCI",
                "Payment Thank You", -card_total, "Credit Card Payment")

        add(month_start.replace(day=min(28, days_in_month)), ACCOUNT_ALEX, PERSON_A,
            "MONTHLY ACCOUNT FEE", "Monthly Account Fee", -11.95, "Bank Fees")
        add(month_start.replace(day=min(11, days_in_month)), ACCOUNT_ALEX, PERSON_A,
            "ATM WITHDRAWAL 100 QUEEN ST W", "Atm Withdrawal", -60, "Cash Withdrawal")

        # A donation and a dental visit every other month.
        if month_start.month % 2 == 0:
            add(month_start.replace(day=min(22, days_in_month)), ACCOUNT_SAM, PERSON_B,
                "DONATION UNITED WAY MISC PAYMENT", "United Way", -50, "Donations")
            add(month_start.replace(day=min(9, days_in_month)), ACCOUNT_JOINT, "Joint",
                "BROADVIEW DENTAL CENTRE", "Broadview Dental Centre", -absish(rng, 120, 340),
                "Dental")

    out.sort(key=lambda txn: (txn.when, txn.account, txn.description))
    return out


def absish(rng: random.Random, low: float, high: float) -> float:
    return round(rng.uniform(low, high), 2)


def _pick_account(rng: random.Random, category: str):
    roll = rng.random()
    if category in ("Groceries", "Public Transit", "Fuel"):
        if roll < 0.45:
            return ACCOUNT_JOINT, "Joint"
        if roll < 0.75:
            return ACCOUNT_CARD, "Joint"
        return (ACCOUNT_ALEX, PERSON_A) if roll < 0.9 else (ACCOUNT_SAM, PERSON_B)
    if roll < 0.35:
        return ACCOUNT_CARD, "Joint"
    if roll < 0.6:
        return ACCOUNT_ALEX, PERSON_A
    if roll < 0.85:
        return ACCOUNT_SAM, PERSON_B
    return ACCOUNT_JOINT, "Joint"


def _days_in_month(when: date) -> int:
    if when.month == 12:
        return 31
    return (date(when.year, when.month + 1, 1) - timedelta(days=1)).day


# ---------------------------------------------------------------------------
# CSV renderings, one per account, each in its bank's real export shape.
# ---------------------------------------------------------------------------


def csv_files(records: List[Txn], today: date | None = None) -> Dict[str, str]:
    today = today or date.today()
    return {
        "rbc-chequing-alex.csv": _rbc(records, ACCOUNT_ALEX),
        "tangerine-chequing-sam.csv": _tangerine(records, ACCOUNT_SAM),
        "bmo-joint-chequing.csv": _bmo(records, ACCOUNT_JOINT, today),
        "amex-cobalt-joint.csv": _amex(records, ACCOUNT_CARD),
    }


def _for(records: List[Txn], account: str) -> List[Txn]:
    return [txn for txn in records if txn.account == account]


def _quote(text: str) -> str:
    if any(char in text for char in ',"\n'):
        return '"' + text.replace('"', '""') + '"'
    return text


def _rbc(records: List[Txn], account: str) -> str:
    lines = [
        '"Account Type","Account Number","Transaction Date","Cheque Number",'
        '"Description 1","Description 2","CAD$","USD$"'
    ]
    for txn in _for(records, account):
        first, second = _split_description(txn.description)
        lines.append(
            ",".join(
                [
                    "Chequing",
                    "04421-1234567",
                    txn.when.strftime("%m/%d/%Y"),
                    "",
                    _quote(first),
                    _quote(second),
                    f"{txn.amount:.2f}",
                    "",
                ]
            )
        )
    return "\n".join(lines) + "\n"


def _tangerine(records: List[Txn], account: str) -> str:
    lines = ["Date,Transaction,Name,Memo,Amount"]
    for txn in _for(records, account):
        kind = "CREDIT" if txn.amount > 0 else "DEBIT"
        lines.append(
            ",".join(
                [
                    txn.when.strftime("%m/%d/%Y"),
                    kind,
                    _quote(txn.description),
                    "",
                    f"{txn.amount:.2f}",
                ]
            )
        )
    return "\n".join(lines) + "\n"


def _bmo(records: List[Txn], account: str, today: date) -> str:
    lines = [
        "Following data is valid as of " + today.strftime("%Y%m%d") + ".",
        "",
        "Transaction history is available for the last 90 days.",
        "",
        "First Bank Card,Transaction Type,Date Posted,"
        "Transaction Amount,Description",
    ]
    for txn in _for(records, account):
        kind = "CR" if txn.amount > 0 else "DR"
        lines.append(
            ",".join(
                [
                    "'5191230000000000'",
                    kind,
                    txn.when.strftime("%Y%m%d"),
                    f"{txn.amount:.2f}",
                    _quote(txn.description),
                ]
            )
        )
    return "\n".join(lines) + "\n"


def _amex(records: List[Txn], account: str) -> str:
    lines = ["Date,Description,Amount,Extended Details"]
    for txn in _for(records, account):
        # Amex writes purchases as positive amounts.
        lines.append(
            ",".join(
                [
                    txn.when.strftime("%Y-%m-%d"),
                    _quote(txn.description),
                    f"{-txn.amount:.2f}",
                    "",
                ]
            )
        )
    return "\n".join(lines) + "\n"


def _split_description(description: str):
    """RBC splits long descriptions over two columns."""
    if len(description) <= 30:
        return description, ""
    cut = description.rfind(" ", 0, 30)
    if cut < 0:
        cut = 30
    return description[:cut], description[cut + 1 :]
