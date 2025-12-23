"""
MCP Finance Data Server - Streamable HTTP Transport

This server provides company earnings data tools via the Model Context Protocol (MCP).
Designed for deployment on Azure Container Apps with Foundry Agent integration.
"""

import os
import json
import logging
from contextlib import asynccontextmanager
from typing import AsyncIterator

import httpx
from dotenv import load_dotenv
from fastmcp import FastMCP
from starlette.requests import Request
from starlette.responses import JSONResponse

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("mcp-finance-server")

# Load environment variables
load_dotenv()

# Initialize MCP server
mcp = FastMCP(
    "MCP Finance Data Server",
    dependencies=["httpx", "python-dotenv"],
)

# Valid company symbols
VALID_SYMBOLS = ["MSFT", "TSLA", "NVDA"]


def get_api_key() -> str:
    """Get Alpha Vantage API key from environment."""
    api_key = os.getenv("ALPHAVANTAGE_API_KEY")
    if not api_key:
        raise ValueError("ALPHAVANTAGE_API_KEY environment variable is required")
    return api_key


def map_month_to_quarter(month: int) -> int | None:
    """Map calendar month to fiscal quarter (approximate for tech companies)."""
    if month in [9, 10]:
        return 1  # Q1 (Oct-Dec)
    elif month in [12, 1]:
        return 2  # Q2 (Jan-Mar)
    elif month in [3, 4]:
        return 3  # Q3 (Apr-Jun)
    elif month in [6, 7]:
        return 4  # Q4 (Jul-Sep)
    return None


async def fetch_alpha_vantage(function: str, symbol: str) -> dict:
    """Fetch data from Alpha Vantage API."""
    api_key = get_api_key()
    url = "https://www.alphavantage.co/query"
    params = {
        "function": function,
        "symbol": symbol,
        "apikey": api_key,
    }
    
    async with httpx.AsyncClient() as client:
        response = await client.get(url, params=params, timeout=15.0)
        response.raise_for_status()
        return response.json()


def find_quarterly_report(reports: list, fiscal_year: int, fiscal_quarter: int) -> dict | None:
    """Find a specific quarterly report by fiscal year and quarter."""
    for report in reports:
        fiscal_date = report.get("fiscalDateEnding", "")
        if not fiscal_date or len(fiscal_date) < 7:
            continue
            
        report_year = int(fiscal_date[:4])
        report_month = int(fiscal_date[5:7])
        report_quarter = map_month_to_quarter(report_month)
        
        if report_quarter and report_year == fiscal_year and report_quarter == fiscal_quarter:
            return report
    
    # Fallback: return most recent if no exact match
    return reports[0] if reports else None


@mcp.tool()
async def get_company_revenue(
    company_symbol: str,
    fiscal_year: int = 2024,
    fiscal_quarter: int = 4,
) -> str:
    """
    Get total revenue in USD millions for a company's quarterly earnings report.
    
    Args:
        company_symbol: Company ticker symbol (MSFT, TSLA, or NVDA)
        fiscal_year: Fiscal year (e.g., 2024)
        fiscal_quarter: Fiscal quarter (1, 2, 3, or 4)
    
    Returns:
        JSON string with revenue data including total_revenue_usd_millions
    """
    symbol = company_symbol.upper()
    
    # Validate inputs
    if symbol not in VALID_SYMBOLS:
        return json.dumps({
            "status": "error",
            "message": f"Invalid company symbol. Supported: {', '.join(VALID_SYMBOLS)}"
        }, indent=2)
    
    if fiscal_quarter not in [1, 2, 3, 4]:
        return json.dumps({
            "status": "error",
            "message": "Fiscal quarter must be 1, 2, 3, or 4"
        }, indent=2)
    
    try:
        data = await fetch_alpha_vantage("INCOME_STATEMENT", symbol)
        quarterly_reports = data.get("quarterlyReports", [])
        
        if not quarterly_reports:
            return json.dumps({
                "status": "error",
                "message": "No quarterly data available"
            }, indent=2)
        
        target_period = find_quarterly_report(quarterly_reports, fiscal_year, fiscal_quarter)
        
        if not target_period:
            return json.dumps({
                "status": "error",
                "message": f"No data found for {symbol} FY{fiscal_year} Q{fiscal_quarter}"
            }, indent=2)
        
        # Extract revenue
        revenue_str = target_period.get("totalRevenue", "0")
        revenue_millions = round(float(revenue_str) / 1_000_000, 2) if revenue_str and revenue_str != "None" else 0
        
        result = {
            "company": symbol,
            "fiscal_year": fiscal_year,
            "fiscal_quarter": fiscal_quarter,
            "fiscal_date_ending": target_period.get("fiscalDateEnding"),
            "total_revenue_usd_millions": revenue_millions,
            "currency": "USD",
            "data_source": "Alpha Vantage",
        }
        
        logger.info(f"Revenue for {symbol} FY{fiscal_year} Q{fiscal_quarter}: ${revenue_millions}M")
        return json.dumps(result, indent=2)
        
    except httpx.HTTPError as e:
        logger.error(f"HTTP error fetching revenue: {e}")
        return json.dumps({
            "status": "error",
            "message": f"HTTP error: {str(e)}"
        }, indent=2)
    except Exception as e:
        logger.error(f"Error fetching revenue: {e}")
        return json.dumps({
            "status": "error",
            "message": str(e)
        }, indent=2)


@mcp.tool()
async def get_company_free_cash_flow(
    company_symbol: str,
    fiscal_year: int = 2024,
    fiscal_quarter: int = 4,
) -> str:
    """
    Get free cash flow in USD millions for a company's quarterly earnings report.
    
    Free Cash Flow = Operating Cash Flow - Capital Expenditures
    
    Args:
        company_symbol: Company ticker symbol (MSFT, TSLA, or NVDA)
        fiscal_year: Fiscal year (e.g., 2024)
        fiscal_quarter: Fiscal quarter (1, 2, 3, or 4)
    
    Returns:
        JSON string with free cash flow data
    """
    symbol = company_symbol.upper()
    
    # Validate inputs
    if symbol not in VALID_SYMBOLS:
        return json.dumps({
            "status": "error",
            "message": f"Invalid company symbol. Supported: {', '.join(VALID_SYMBOLS)}"
        }, indent=2)
    
    if fiscal_quarter not in [1, 2, 3, 4]:
        return json.dumps({
            "status": "error",
            "message": "Fiscal quarter must be 1, 2, 3, or 4"
        }, indent=2)
    
    try:
        data = await fetch_alpha_vantage("CASH_FLOW", symbol)
        quarterly_reports = data.get("quarterlyReports", [])
        
        if not quarterly_reports:
            return json.dumps({
                "status": "error",
                "message": "No quarterly cash flow data available"
            }, indent=2)
        
        target_period = find_quarterly_report(quarterly_reports, fiscal_year, fiscal_quarter)
        
        if not target_period:
            return json.dumps({
                "status": "error",
                "message": f"No data found for {symbol} FY{fiscal_year} Q{fiscal_quarter}"
            }, indent=2)
        
        # Calculate Free Cash Flow
        op_cashflow_str = target_period.get("operatingCashflow", "0")
        capex_str = target_period.get("capitalExpenditures", "0")
        
        op_cashflow = float(op_cashflow_str) if op_cashflow_str and op_cashflow_str != "None" else 0
        capex = float(capex_str) if capex_str and capex_str != "None" else 0
        
        free_cash_flow = op_cashflow - abs(capex)
        fcf_millions = round(free_cash_flow / 1_000_000, 2)
        
        result = {
            "company": symbol,
            "fiscal_year": fiscal_year,
            "fiscal_quarter": fiscal_quarter,
            "fiscal_date_ending": target_period.get("fiscalDateEnding"),
            "free_cash_flow_usd_millions": fcf_millions,
            "operating_cash_flow_usd_millions": round(op_cashflow / 1_000_000, 2),
            "capital_expenditures_usd_millions": round(abs(capex) / 1_000_000, 2),
            "currency": "USD",
            "data_source": "Alpha Vantage",
            "calculation": "Operating Cash Flow - Capital Expenditures"
        }
        
        logger.info(f"Free Cash Flow for {symbol} FY{fiscal_year} Q{fiscal_quarter}: ${fcf_millions}M")
        return json.dumps(result, indent=2)
        
    except httpx.HTTPError as e:
        logger.error(f"HTTP error fetching cash flow: {e}")
        return json.dumps({
            "status": "error",
            "message": f"HTTP error: {str(e)}"
        }, indent=2)
    except Exception as e:
        logger.error(f"Error fetching cash flow: {e}")
        return json.dumps({
            "status": "error",
            "message": str(e)
        }, indent=2)


# Health check endpoint (for Container Apps)
@mcp.custom_route("/health", methods=["GET"])
async def health_check(request: Request) -> JSONResponse:
    """Health check endpoint for Container Apps."""
    return JSONResponse({"status": "healthy", "service": "mcp-finance-server"})


if __name__ == "__main__":
    port = int(os.getenv("PORT", "3000"))
    host = os.getenv("HOST", "0.0.0.0")
    
    logger.info(f"Starting MCP Finance Server on {host}:{port}")
    logger.info(f"MCP endpoint: http://{host}:{port}/mcp")
    
    # Run with HTTP transport (streamable HTTP)
    mcp.run(transport="http", host=host, port=port)
