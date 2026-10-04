"""Parse Venezuelan and US-style receipt amounts without locale ambiguity."""
from __future__ import annotations

import math
import re
from decimal import Decimal, InvalidOperation


_ERROR = {
    "es": "Usa un importe válido, por ejemplo 749,50 o 749.50 Bs.",
    "en": "Enter a valid amount, for example 749,50 or 749.50 Bs.",
}

_RATE_ERROR = {
    "es": "Usa una tasa válida con hasta 8 decimales, por ejemplo 871,36890000 o 871.36890000.",
    "en": "Enter a valid rate with up to 8 decimals, for example 871,36890000 or 871.36890000.",
}


def _language(language: str | None) -> str:
    return "en" if (language or "").lower().startswith("en") else "es"


def _valid_integer_groups(groups: list[str]) -> bool:
    if not groups or not all(group.isdigit() for group in groups):
        return False
    if len(groups) == 1:
        return True
    return 1 <= len(groups[0]) <= 3 and all(len(group) == 3 for group in groups[1:])


def parse_monto_usuario(value: object, language: str = "es") -> float:
    """Parse a nonnegative amount accepting comma or dot decimals.

    Supports ``749,50``, ``749.50``, ``1.250,50`` and ``1,250.50``.
    A single separator followed by three digits is treated as a thousands
    separator, matching common Venezuelan and US display conventions.
    """
    if value is None:
        return 0.0
    if isinstance(value, bool):
        raise ValueError(_ERROR[_language(language)])
    if isinstance(value, (int, float)):
        amount = float(value)
        if not math.isfinite(amount) or amount < 0:
            raise ValueError(_ERROR[_language(language)])
        return round(amount, 2)

    raw = str(value).strip()
    raw = re.sub(r"(?i)\s*bs\.?\s*$", "", raw).strip()
    raw = raw.replace(" ", "")
    if not raw:
        return 0.0
    if not re.fullmatch(r"\d+(?:[.,]\d+)*", raw):
        raise ValueError(_ERROR[_language(language)])

    comma_count = raw.count(",")
    dot_count = raw.count(".")
    integer = raw
    fraction = ""

    if comma_count and dot_count:
        decimal_separator = "," if raw.rfind(",") > raw.rfind(".") else "."
        group_separator = "." if decimal_separator == "," else ","
        integer_part, fraction = raw.rsplit(decimal_separator, 1)
        groups = integer_part.split(group_separator)
        if not _valid_integer_groups(groups) or not 1 <= len(fraction) <= 2:
            raise ValueError(_ERROR[_language(language)])
        integer = "".join(groups)
    elif comma_count or dot_count:
        separator = "," if comma_count else "."
        parts = raw.split(separator)
        if len(parts) == 2:
            before, after = parts
            if len(after) == 3 and 1 <= len(before) <= 3:
                integer = before + after
            elif 1 <= len(after) <= 2:
                integer, fraction = before, after
            else:
                raise ValueError(_ERROR[_language(language)])
        elif _valid_integer_groups(parts):
            integer = "".join(parts)
        elif 1 <= len(parts[-1]) <= 2 and _valid_integer_groups(parts[:-1]):
            integer = "".join(parts[:-1])
            fraction = parts[-1]
        else:
            raise ValueError(_ERROR[_language(language)])

    try:
        amount = Decimal(integer + ("." + fraction if fraction else ""))
    except InvalidOperation:
        raise ValueError(_ERROR[_language(language)]) from None
    if not amount.is_finite() or amount < 0:
        raise ValueError(_ERROR[_language(language)])
    return float(amount.quantize(Decimal("0.01")))


def parse_tasa_usuario(value: object, language: str = "es") -> Decimal:
    """Parse a BCV/manual rate with up to eight decimal places.

    Decimal comma and decimal point are accepted. When both separators are
    present, the last one is the decimal separator and the other must form
    valid thousands groups.
    """
    error = _RATE_ERROR[_language(language)]
    if value is None:
        return Decimal("0")
    if isinstance(value, bool):
        raise ValueError(error)
    if isinstance(value, Decimal):
        result = value
    elif isinstance(value, (int, float)):
        if not math.isfinite(float(value)):
            raise ValueError(error)
        result = Decimal(str(value))
    else:
        raw = str(value).strip().replace(" ", "")
        if not raw:
            return Decimal("0")
        if not re.fullmatch(r"\d+(?:[.,]\d+)*", raw):
            raise ValueError(error)
        if "," in raw and "." in raw:
            decimal_separator = "," if raw.rfind(",") > raw.rfind(".") else "."
            group_separator = "." if decimal_separator == "," else ","
            integer_part, fraction = raw.rsplit(decimal_separator, 1)
            groups = integer_part.split(group_separator)
            if not _valid_integer_groups(groups) or not 1 <= len(fraction) <= 8:
                raise ValueError(error)
            normalized = "".join(groups) + "." + fraction
        elif "," in raw or "." in raw:
            parts = re.split(r"[,.]", raw)
            if len(parts) == 2 and 1 <= len(parts[1]) <= 8:
                normalized = parts[0] + "." + parts[1]
            elif _valid_integer_groups(parts):
                normalized = "".join(parts)
            else:
                raise ValueError(error)
        else:
            normalized = raw
        try:
            result = Decimal(normalized)
        except InvalidOperation:
            raise ValueError(error) from None

    if not result.is_finite() or result < 0 or result.as_tuple().exponent < -8:
        raise ValueError(error)
    return result
