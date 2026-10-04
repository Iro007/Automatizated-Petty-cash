"""Read the official USD reference rate and value date from the BCV page."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timezone
from html.parser import HTMLParser
import re
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from amount_utils import parse_tasa_usuario

BCV_URL = "https://www.bcv.org.ve/estadisticas/tipo-cambio-de-referencia-smc"
_SPANISH_MONTHS = {
    "enero": 1,
    "febrero": 2,
    "marzo": 3,
    "abril": 4,
    "mayo": 5,
    "junio": 6,
    "julio": 7,
    "agosto": 8,
    "septiembre": 9,
    "setiembre": 9,
    "octubre": 10,
    "noviembre": 11,
    "diciembre": 12,
}


class BCVRateError(ValueError):
    """Raised when the official page cannot provide a valid USD rate/date."""


@dataclass(frozen=True)
class BCVRate:
    rate: str
    value_date: date
    fetched_at: datetime
    source_url: str = BCV_URL


class _BCVHTMLParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.in_dollar_card = False
        self.capture_rate = False
        self.capture_date = False
        self.rate_parts: list[str] = []
        self.date_parts: list[str] = []
        self.date_content: str | None = None

    def handle_starttag(self, tag: str, attrs) -> None:
        attributes = dict(attrs)
        if tag == "div" and attributes.get("id") == "dolar":
            self.in_dollar_card = True
        if (
            self.in_dollar_card
            and tag == "strong"
            and "strong-tb" in (attributes.get("class", "").split())
            and not self.rate_parts
        ):
            self.capture_rate = True
        if tag == "span" and "date-display-single" in attributes.get("class", "").split():
            self.capture_date = True
            self.date_content = attributes.get("content")

    def handle_endtag(self, tag: str) -> None:
        if tag == "strong":
            self.capture_rate = False
        elif tag == "span":
            self.capture_date = False

    def handle_data(self, data: str) -> None:
        if self.capture_rate:
            self.rate_parts.append(data)
        if self.capture_date:
            self.date_parts.append(data)


def _parse_value_date(content: str | None, visible_text: str) -> date:
    if content:
        try:
            return date.fromisoformat(content[:10])
        except ValueError:
            pass

    match = re.search(r"\b(\d{1,2})\s+([A-Za-zÁÉÍÓÚáéíóú]+)\s+(\d{4})\b", visible_text)
    if not match:
        raise BCVRateError("The BCV page did not contain a recognizable value date.")
    day, month_name, year = match.groups()
    month = _SPANISH_MONTHS.get(month_name.lower())
    if month is None:
        raise BCVRateError("The BCV page contained an unknown Spanish month name.")
    try:
        return date(int(year), month, int(day))
    except ValueError as exc:
        raise BCVRateError("The BCV page contained an invalid value date.") from exc


def parse_bcv_html(html: str, fetched_at: datetime | None = None) -> BCVRate:
    """Parse the USD card and official ``Fecha Valor`` from BCV HTML."""
    parser = _BCVHTMLParser()
    parser.feed(html)
    raw_rate = "".join(parser.rate_parts).strip()
    if not raw_rate:
        raise BCVRateError("The BCV USD rate was not found in the official page.")
    try:
        rate = parse_tasa_usuario(raw_rate, language="es")
    except ValueError as exc:
        raise BCVRateError("The BCV USD rate is not a valid number.") from exc
    if rate <= 0:
        raise BCVRateError("The BCV USD rate must be greater than zero.")

    value_date = _parse_value_date(parser.date_content, " ".join(parser.date_parts))
    timestamp = fetched_at or datetime.now(timezone.utc)
    if timestamp.tzinfo is None:
        timestamp = timestamp.replace(tzinfo=timezone.utc)
    return BCVRate(
        rate=format(rate, ".8f"),
        value_date=value_date,
        fetched_at=timestamp,
    )


def fetch_bcv_rate(timeout: float = 12.0) -> BCVRate:
    """Fetch the current official BCV page with a bounded request timeout."""
    request = Request(
        BCV_URL,
        headers={"User-Agent": "CajaChicaPro/1.0 (+BCV public rate lookup)"},
    )
    try:
        with urlopen(request, timeout=timeout) as response:
            status = getattr(response, "status", 200)
            if status != 200:
                raise BCVRateError(f"The BCV page returned HTTP {status}.")
            charset = response.headers.get_content_charset() or "utf-8"
            html = response.read().decode(charset, errors="replace")
    except HTTPError as exc:
        raise BCVRateError(f"The BCV page returned HTTP {exc.code}.") from exc
    except (TimeoutError, URLError, OSError) as exc:
        raise BCVRateError("The BCV page could not be reached.") from exc
    return parse_bcv_html(html)
