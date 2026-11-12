"""Display-only hints that explain how a vendor's month-end balance was offset.

Nothing here changes a reconciliation result. Three kinds of information are
kept strictly apart, because only the first is a confirmed pairing:

* confirmed one-to-one pairs come from the reconciliation itself (not here);
* the vendor arithmetic (opening + credits - debits = closing) is a fact;
* a "month group" is only a *candidate*: payments and claims whose descriptions
  name the same single month (e.g. "지급(7월분)" and "7월 위탁물류비").
"""
import re
from dataclasses import dataclass
from decimal import Decimal

from .normalization import normalize_text

_MONTH=re.compile(r"(?<!\d)(1[0-2]|0?[1-9])\s*월")


def month_token(text):
    """The one month a description names, or None.

    A description that names no month, or several different months ("9월 임차료
    및 8월 관리비"), is not grouped: guessing would invent a link.
    """
    months={int(m.group(1)) for m in _MONTH.finditer(text or "")}
    return months.pop() if len(months)==1 else None


def differs_only_by_month(left, right):
    """True when two descriptions are the same item and only the month number differs.

    "7월 위탁물류비용_GS팀" and "8월 위탁물류비용_GS팀" -> True. Anything else that
    differs (another team, another word) -> False; identical text -> False.
    """
    a=normalize_text(left)
    b=normalize_text(right)
    return (a!=b and _MONTH.search(a) is not None
            and _MONTH.sub("#월",a)==_MONTH.sub("#월",b))


def offsetting_pairs(rows):
    """Rows that offset each other one-to-one: the same item (month ignored), equal
    amounts of opposite sign, and nothing else that could be the partner.

    rows: iterable of (key, description, signed amount). Returns {key: partner key}
    for both sides of every unambiguous pair; ambiguous rows are left out.
    """
    buckets={}
    for key,description,amount in rows:
        amount=Decimal(amount)
        if amount==0:
            continue
        item=_MONTH.sub("#월",normalize_text(description))
        buckets.setdefault((item,abs(amount)),([],[]))[0 if amount>0 else 1].append(key)
    pairs={}
    for positive,negative in buckets.values():
        if len(positive)==1 and len(negative)==1:
            pairs[positive[0]]=negative[0]
            pairs[negative[0]]=positive[0]
    return pairs


@dataclass(frozen=True)
class MonthGroup:
    month: int
    claims: Decimal      # signed sum of the opening rows of that month (adjustments included)
    payments: Decimal    # sum of the current-month payments naming that month (positive)

    @property
    def remaining(self):
        return self.claims-self.payments


@dataclass(frozen=True)
class VendorNote:
    """What the vendor subtotal row explains."""
    opening: Decimal
    credit: Decimal
    debit: Decimal       # positive amount paid
    closing: Decimal     # the closing balance shown on the statement
    expected: Decimal    # opening + credit - debit
    state: str           # "ok" | "mismatch" | "hold" (inputs incomplete)
    review: bool         # some row of the vendor still needs a person
    groups: tuple = ()


@dataclass(frozen=True)
class LinkNote:
    """"This row is offset against those rows": the wording, resolved to row numbers when written."""
    head: str            # e.g. "자동 대사", "순잔액 대사 · 7월 묶음"
    noun: str            # what the partner rows are, e.g. "지급", "7월 청구"; "" for a confirmed pair
    partners: tuple      # ids of the partner StatementRow objects
    confirmed: bool      # True for a one-to-one pair (equal and opposite amounts, same item)
    suffix: str = ""     # e.g. " (월만 다름)" when only the month in the description differs


def month_groups(claims, payments):
    """Group by month the months that appear on BOTH sides.

    claims: iterable of (description, signed amount); payments: iterable of
    (description, positive amount). Returns {month: MonthGroup}.
    """
    claim_sum={}
    for description,amount in claims:
        month=month_token(description)
        if month is not None:
            claim_sum[month]=claim_sum.get(month,Decimal(0))+Decimal(amount)
    payment_sum={}
    for description,amount in payments:
        month=month_token(description)
        if month is not None:
            payment_sum[month]=payment_sum.get(month,Decimal(0))+Decimal(amount)
    return {
        month:MonthGroup(month,claim_sum[month],payment_sum[month])
        for month in sorted(set(claim_sum)&set(payment_sum))
    }
