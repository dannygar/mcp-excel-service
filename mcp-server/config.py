"""
Configuration module for MCP Excel Service.

Contains trade tracker settings and strategy name mappings.
"""

import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# =============================================================================
# Trade Tracker Configuration (from environment variables)
# =============================================================================
TRADE_TRACKER_URL = os.getenv(
    "TRADE_TRACKER_URL",
    "https://mngenvmcap046191.sharepoint.com/Shared%20Documents/Forms/AllItems.aspx"
)
TRADE_TRACKER_FILE = os.getenv(
    "TRADE_TRACKER_FILE",
    "2026 GY Capital Group LLC Trade Tracker.xlsx"
)

# =============================================================================
# Strategy Name Mapping (Full Name → Excel Short Code)
# =============================================================================
# Maps verbose strategy names to their abbreviated codes used in the Excel tracker
# Includes common variations and abbreviations for fuzzy matching
STRATEGY_MAPPING = {
    # Case-insensitive lookup - keys are lowercase
    # Covered Call variations
    "covered call": "CC",
    "cc": "CC",
    
    # Naked Put variations
    "naked put": "NP",
    "np": "NP",
    
    # Cash-Secured Put variations
    "cash-secured put": "CSP",
    "cash secured put": "CSP",
    "csp": "CSP",
    
    # Iron Condor variations
    "iron condor": "IC",
    "ic": "IC",
    "condor": "IC",
    
    # Iron Butterfly variations
    "iron butterfly": "IB",
    "ib": "IB",
    "butterfly": "IB",  # Default butterfly to IB
    
    # Vertical Put Credit Spread variations
    "vertical put credit spread": "VPCS",
    "put credit spread": "VPCS",
    "vertical put credit": "VPCS",
    "vertical put spread": "VPCS",
    "put vertical credit": "VPCS",
    "put vertical": "VPCS",
    "vertical put": "VPCS",
    "pcs": "VPCS",
    "vpcs": "VPCS",
    
    # Vertical Put Debit Spread variations
    "vertical put debit spread": "VPDS",
    "put debit spread": "VPDS",
    "vertical put debit": "VPDS",
    "pds": "VPDS",
    "vpds": "VPDS",
    
    # Vertical Call Credit Spread variations
    "vertical call credit spread": "VCCS",
    "call credit spread": "VCCS",
    "vertical call credit": "VCCS",
    "vertical call spread": "VCCS",
    "call vertical credit": "VCCS",
    "call vertical": "VCCS",
    "vertical call": "VCCS",
    "ccs": "VCCS",
    "vccs": "VCCS",
    
    # Straddle variations
    "short straddle": "Straddle",
    "straddle": "Straddle",
    
    # Strangle variations
    "short strangle": "Strangle",
    "strangle": "Strangle",
    
    # Jade Lizard variations
    "jade lizard": "JadeLizard",
    "jadelizard": "JadeLizard",
    "jade": "JadeLizard",
    
    # Reverse Jade Lizard variations
    "reverse jade lizard": "RJade",
    "reverse jade": "RJade",
    "rjade": "RJade",
    
    # Zebra variations
    "zero extrinsic back ratio": "Zebra",
    "zebra": "Zebra",
    
    # 1-1-1 variations
    "lt1-1-1": "1-1-1",
    "1-1-1": "1-1-1",
    "111": "1-1-1",
    
    # 1-1-2 variations
    "lt1-1-2": "1-1-2",
    "1-1-2": "1-1-2",
    "112": "1-1-2",
    
    # Rolling Diagonal Puts variations
    "rolling diagonal puts": "RDP",
    "rolling diagonal": "RDP",
    "diagonal puts": "RDP",
    "rdp": "RDP",
    
    # Short LEAPS variations
    "short leaps": "LEAPPut",
    "leap put": "LEAPPut",
    "leaps put": "LEAPPut",
    "leapput": "LEAPPut",
    
    # LEAPS Call Spread variations
    "leaps call spread": "LeapCS",
    "leap call spread": "LeapCS",
    "leapcs": "LeapCS",
    
    # Put Butterfly variations
    "put butterfly": "PButterfly",
    "pbutterfly": "PButterfly",
    
    # Call Butterfly variations
    "call butterfly": "CButterfly",
    "cbutterfly": "CButterfly",
    
    # VIX variations
    "vix uptrend": "VIX",
    "vix": "VIX",
}

# =============================================================================
# Keyword-based Fallback Patterns for Fuzzy Matching
# =============================================================================
# Order matters - more specific patterns should come first
STRATEGY_KEYWORDS = [
    # Most specific patterns first
    (["put", "credit", "spread"], "VPCS"),
    (["put", "debit", "spread"], "VPDS"),
    (["call", "credit", "spread"], "VCCS"),
    (["call", "debit", "spread"], "VCDS"),
    (["iron", "condor"], "IC"),
    (["iron", "butterfly"], "IB"),
    (["put", "butterfly"], "PButterfly"),
    (["call", "butterfly"], "CButterfly"),
    (["jade", "lizard"], "JadeLizard"),
    (["reverse", "jade"], "RJade"),
    (["rolling", "diagonal"], "RDP"),
    (["diagonal", "put"], "RDP"),
    (["cash", "secured"], "CSP"),
    (["covered", "call"], "CC"),
    (["naked", "put"], "NP"),
    (["leap", "call"], "LeapCS"),
    (["leap", "put"], "LEAPPut"),
    # Vertical spreads (put/call + vertical)
    (["put", "vertical"], "VPCS"),
    (["vertical", "put"], "VPCS"),
    (["call", "vertical"], "VCCS"),
    (["vertical", "call"], "VCCS"),
    # Single keyword patterns (less specific)
    (["straddle"], "Straddle"),
    (["strangle"], "Strangle"),
    (["condor"], "IC"),
    (["zebra"], "Zebra"),
]


def map_strategy_name(strategy: str) -> str:
    """
    Map a strategy name to its Excel short code.
    Uses exact matching first, then keyword-based fuzzy matching.
    If already a short code or not found, returns the original value.
    
    Args:
        strategy: The strategy name to map (e.g., "Iron Condor", "Put Vertical")
    
    Returns:
        The Excel short code (e.g., "IC", "VPCS") or the original value if no match
    """
    if not strategy:
        return strategy
    
    # Normalize input
    lookup_key = strategy.lower().strip()
    
    # Check if it's already a known short code (case-insensitive)
    short_codes = set(STRATEGY_MAPPING.values())
    for code in short_codes:
        if lookup_key == code.lower():
            return code  # Return the properly-cased short code
    
    # Try exact match in mapping
    if lookup_key in STRATEGY_MAPPING:
        return STRATEGY_MAPPING[lookup_key]
    
    # Try keyword-based fuzzy matching
    for keywords, short_code in STRATEGY_KEYWORDS:
        if all(keyword in lookup_key for keyword in keywords):
            return short_code
    
    # No match found - return original
    return strategy
