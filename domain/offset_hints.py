"""Display-only hints that explain how a vendor's month-end balance was offset.

Nothing here changes a reconciliation result. Three kinds of information are
kept strictly apart, because only the first is a confirmed pairing:

* confirmed one-to-one pairs come from the reconciliation itself (not here);
* the vendor arithmetic (opening + credits - debits = closing) is a fact;
* a "month group" is only a *candidate*: payments and claims whose descriptions
  name the same single month (e.g. "지급(7월분)" and "7월 위탁물류비").
"""
import itertools
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


# Amount-only matching of different items ("A + B = C") is only trusted when a
# coincidence is very unlikely: few rows, small sets, and a sum that is not a round number.
MAX_MATCH_ROWS=20
MAX_MATCH_SIZE=4
MIN_SIGNIFICANT_DIGITS=5


def significant_digits(amount):
    """Digits of an amount without trailing zeros: 100,000,000 -> 1, 807,567,651 -> 9."""
    amount=abs(Decimal(amount))
    text=str(int(amount)) if amount==amount.to_integral_value() else format(amount.normalize(),"f").replace(".","")
    return len(text.rstrip("0")) or 1


def _unique_amount_matches(pool):
    """Zero-sum sets of DIFFERENT items found by amount alone, only when the answer is unique.

    pool: [(key, signed amount)]. For each row, the rows of the opposite sign (at most
    MAX_MATCH_SIZE of them) that add up to its amount must be exactly one combination;
    a row that two different candidate sets claim is dropped from both.
    """
    if len(pool)>MAX_MATCH_ROWS:
        return []
    positive=[(k,a) for k,a in pool if a>0]
    negative=[(k,-a) for k,a in pool if a<0]
    candidates=set()
    for anchors,others in ((negative,positive),(positive,negative)):
        for key,amount in anchors:
            if significant_digits(amount)<MIN_SIGNIFICANT_DIGITS:
                continue
            combos=[c for size in range(1,MAX_MATCH_SIZE+1)
                    for c in itertools.combinations(others,size)
                    if sum(x[1] for x in c)==amount]
            if len(combos)==1:
                candidates.add(frozenset([key,*[x[0] for x in combos[0]]]))
    return [s for s in candidates if all(s==t or not (s&t) for t in candidates)]


@dataclass(frozen=True)
class OffsetSet:
    """Rows that offset each other: a one-to-one pair, or a group whose amounts sum to 0."""
    keys: tuple
    kind: str            # "pair" | "group" | "mixed" (different items, found by amount only)


def offset_sets(rows):
    """Find rows of one vendor that offset each other, strictest evidence first.

    rows: iterable of (key, description, signed amount). Returns disjoint OffsetSets.
    1. pair: the same description, equal amounts of opposite sign, and nobody else
       that fits (a single positive and a single negative);
    2. pair: the same item with the month ignored, same rule;
    3. group: rows of the same item (month ignored) that contain both signs and add up
       to exactly 0 - e.g. a claim of 807 paid as 500 + 307, or 100 + 200 paid as 300.
    4. mixed: different items that add up to 0 by amount alone, only when unique and
       the amount is unlikely to be a coincidence (see MIN_SIGNIFICANT_DIGITS).
    Anything ambiguous or not exactly zero is left out: no guessing.
    """
    remaining=[]
    for key,description,amount in rows:
        amount=Decimal(amount)
        if amount!=0:
            text=normalize_text(description)
            remaining.append((key,text,_MONTH.sub("#월",text),amount))
    found=[]

    def take(group_key, kind, exact_pairs):
        nonlocal remaining
        buckets={}
        for row in remaining:
            buckets.setdefault(group_key(row),[]).append(row)
        used=set()
        for rows_in in buckets.values():
            positive=[r for r in rows_in if r[3]>0]
            negative=[r for r in rows_in if r[3]<0]
            if exact_pairs:
                ok=len(positive)==1 and len(negative)==1
            else:
                ok=bool(positive) and bool(negative) and sum(r[3] for r in rows_in)==0
            if ok:
                found.append(OffsetSet(tuple(r[0] for r in rows_in),kind))
                used.update(r[0] for r in rows_in)
        remaining=[r for r in remaining if r[0] not in used]

    take(lambda r:(r[1],abs(r[3])),"pair",True)
    take(lambda r:(r[2],abs(r[3])),"pair",True)
    take(lambda r:r[2],"group",False)
    # Different items that only add up by amount ("A + B = C"): a candidate, never more.
    for keys in _unique_amount_matches([(r[0],r[3]) for r in remaining]):
        found.append(OffsetSet(tuple(sorted(keys,key=str)),"mixed"))
    return found


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
    hold_cause: str = ""   # why the check is on hold, in words


@dataclass(frozen=True)
class LinkNote:
    """"This row is offset against those rows": the wording, resolved to row numbers when written."""
    head: str            # e.g. "자동 대사", "순잔액 대사 · 7월 묶음"
    noun: str            # what the partner rows are, e.g. "지급", "7월 청구"; "" for a confirmed pair
    partners: tuple      # ids of the partner StatementRow objects
    confirmed: bool      # True for a one-to-one pair (equal and opposite amounts, same item)
    suffix: str = ""     # e.g. " (월만 다름)" when only the month in the description differs
    together: bool = False   # a group (several rows that add up to 0), not a pair
    append: bool = False     # add under the engine's own status line instead of replacing it
    candidate: bool = False  # found by amount alone / engine pair with a different text: "후보"


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
