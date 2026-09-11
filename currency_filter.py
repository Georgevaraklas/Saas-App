"""
Currency display filter.

Formats a Decimal (or any numeric) amount as "1,234.56 $" / "1,234.56 €" etc.
No exchange-rate conversion happens here — the stored amount is displayed
as-is, just labeled with the symbol/code for the account's chosen currency.
"""

from decimal import Decimal, InvalidOperation

# Only currencies with a widely-recognized, unambiguous symbol are listed
# here. Anything not in this map falls back to its 3-letter ISO code
# (e.g. "42.00 PLN"), per your requirement — avoids ambiguous symbols like
# kr (SEK/NOK/DKK all use it) or the several currencies that use "$".
CURRENCY_SYMBOLS = {
    "USD": "$",
    "EUR": "€",
    "GBP": "£",
    "JPY": "¥",
    "CNY": "¥",
    "CHF": "CHF",   # no distinct symbol in common use; kept explicit
    "INR": "₹",
    "KRW": "₩",
    "VND": "₫",
    "TRY": "₺",
    "ILS": "₪",
    "THB": "฿",
    "PHP": "₱",
}


def format_currency(amount, currency_code):
    """
    Format `amount` as a string with 2 decimals, comma thousands separator,
    followed by the currency symbol (or ISO code if no symbol is mapped).

    Examples:
        format_currency(Decimal("1234.5"), "USD")  -> "1,234.50 $"
        format_currency(Decimal("42"), "PLN")       -> "42.00 PLN"
        format_currency(None, "USD")                -> ""
        format_currency("n/a", "USD")                -> "n/a"  (unparseable: show raw value, never blank)
    """
    if amount is None or amount == "":
        return ""

    try:
        amount = Decimal(str(amount))
    except (InvalidOperation, ValueError, TypeError):
        # Don't silently blank the cell if the value isn't a clean number -
        # a table row missing its number is worse and harder to notice
        # than one showing an unformatted raw value.
        return str(amount)

    code = (currency_code or "USD").upper()
    label = CURRENCY_SYMBOLS.get(code, code)

    # Always 2 decimals, comma thousands separator
    formatted_number = f"{amount:,.2f}"

    return f"{formatted_number} {label}"


def register_currency_filter(app):
    """Call once in your app factory: register_currency_filter(app)"""
    app.jinja_env.filters["currency"] = format_currency
