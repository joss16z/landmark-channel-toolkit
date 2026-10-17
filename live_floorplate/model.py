"""Domain records, independent of the unverified getProject response schema."""
from dataclasses import dataclass
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
import re


class FloorplateError(ValueError):
    """Safe failure; UI never renders exception text from upstream libraries."""


def lot_key(value):
    if isinstance(value, bool) or not isinstance(value, (str, int)):
        raise FloorplateError('Invalid lotnumber')
    text = str(value).strip()
    if not re.fullmatch(r'[A-Za-z0-9-]+', text):
        raise FloorplateError('Invalid lotnumber')
    return str(int(text)) if text.isascii() and text.isdecimal() else text.upper()


def number(value, *, price=False):
    if isinstance(value, bool) or value is None:
        raise FloorplateError('Missing number')
    text = str(value).strip()
    text = re.sub(r'^(?:AUD\s*|A)?\$\s*', '', text) if price else re.sub(r'\s*(?:m²|m2|sqm)$', '', text)
    if not re.fullmatch(r'(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?', text):
        raise FloorplateError('Malformed number')
    try:
        result = Decimal(text.replace(',', ''))
    except InvalidOperation:
        raise FloorplateError('Malformed number') from None
    if not result.is_finite() or result < 0 or (price and result <= 0):
        raise FloorplateError('Invalid number')
    if price and result != result.quantize(Decimal('0.01')):
        raise FloorplateError('Fractional-cent price is unsupported')
    return result


def format_price(value):
    amount = number(value, price=True)
    return f'${amount:,.0f}' if amount == amount.to_integral() else f'${amount:,.2f}'


def text(value):
    if not isinstance(value, (str, int)) or isinstance(value, bool) or not str(value).strip():
        raise FloorplateError('Missing identity field')
    result = str(value).strip()
    if len(result) > 100 or any(ord(c) < 32 for c in result):
        raise FloorplateError('Invalid text')
    return result


@dataclass(frozen=True)
class Apartment:
    lot: str
    unit: str
    bedrooms: str
    internal: Decimal
    external: Decimal
    price: Decimal
    bathrooms: str = ''

    @property
    def label(self):
        return f'{self.configuration}\n{self.internal:g} + {self.external:g} m²\nUnit {self.unit} - {format_price(self.price)}'

    @property
    def configuration(self):
        beds=self.bedrooms+' Bed' if re.fullmatch(r'\d+',self.bedrooms) else self.bedrooms
        baths=self.bathrooms if 'bath' in self.bathrooms.lower() else self.bathrooms+' Bath'
        return f'{beds} + {baths}' if self.bathrooms else beds


@dataclass(frozen=True)
class Snapshot:
    project: str
    apartments: tuple[Apartment, ...]
    synced_at: datetime
    source_rows: int
    sample: bool

    def fresh(self, ttl=300, now=None):
        age = ((now or datetime.now(timezone.utc)) - self.synced_at).total_seconds()
        return 0 <= age < ttl


def snapshot(rows, total, project, *, sample=False):
    """Input is an explicit adapter output, never a guessed API response."""
    if type(total) is not int or total < 0 or not isinstance(rows, list) or len(rows) != total:
        raise FloorplateError('Incomplete dataset')
    text(project)
    keys, apartments = set(), []
    for row in rows:
        if not isinstance(row, dict) or row.get('project') != project:
            raise FloorplateError('Project scope mismatch')
        lot = lot_key(row.get('lotnumber'))
        if lot in keys:
            raise FloorplateError('Duplicate project + lotnumber')
        keys.add(lot)
        status = row.get('propertystatus')
        if status not in ('Available', 'Held', 'Pending', 'Sold'):
            raise FloorplateError('Unknown property status')
        if status != 'Available':
            continue
        apartments.append(Apartment(lot, text(row.get('unit')), text(row.get('bedrooms')),
                                    number(row.get('internal')), number(row.get('external')),
                                    number(row.get('price'), price=True),
                                    text(row.get('bathrooms'))))
    return Snapshot(project, tuple(apartments), datetime.now(timezone.utc), total, sample)
