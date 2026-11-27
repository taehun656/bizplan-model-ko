"""Workbook design tokens. RGB strings are shared by Excel, the guide and the verifier."""

FONT = "맑은 고딕"
NUMBER_FONT = "Arial"
ALLOWED_FONTS = frozenset((FONT, NUMBER_FONT))
ORANGE = "FF6600"
SILVER = "C0C0C0"
GRAY = "A6A6A6"
DARK = "595959"
PEACH = "FFB380"
RED = "FF0000"
BLACK = "000000"
WHITE = "FFFFFF"
YELLOW = "FFFF00"
BLUE = "0000FF"

ALLOWED = frozenset((ORANGE, SILVER, GRAY, DARK, PEACH, RED, BLACK, WHITE, YELLOW, BLUE))
TAB_COLORS = (ORANGE, PEACH, DARK, GRAY)
REVENUE_COLORS = (ORANGE, SILVER, GRAY, DARK, PEACH)
BRIDGE_COLORS = (SILVER, ORANGE, RED)        # endpoints, increase, decrease
SCENARIO_COLORS = (SILVER, ORANGE, RED)      # unselected, selected, benchmark
MONTHLY_COLORS = (ORANGE, GRAY, BLACK)       # revenue, EBITDA, operating profit
CASH_COLORS = (ORANGE, DARK, RED)           # cash, debt, cash floor


def rgb(hex_color):
    """Convert an RGB token into the COLORREF integer used by Excel COM."""
    return int.from_bytes(bytes.fromhex(hex_color), "little")
