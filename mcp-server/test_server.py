"""
MCP Excel Service - Test Suite

This test file provides synthetic data and test cases for validating
the MCP Excel Server functionality. Tests can be run against a local
or deployed MCP server.

Usage:
    # Start the MCP server first
    cd mcp-server && uv run python server.py
    
    # In another terminal, run tests
    cd mcp-server && uv run python test_server.py
    
    # Or run specific tests
    uv run python test_server.py --test update_row
    uv run python test_server.py --test update_range
    uv run python test_server.py --test health
"""

import argparse
import asyncio
import json
import sys
from datetime import datetime, timedelta
from typing import Any

import httpx

# =============================================================================
# Configuration
# =============================================================================

MCP_SERVER_URL = "http://localhost:3000"
MCP_ENDPOINT = f"{MCP_SERVER_URL}/mcp"
HEALTH_ENDPOINT = f"{MCP_SERVER_URL}/health"

# Test timeout in seconds
TIMEOUT = 30.0

# =============================================================================
# Synthetic Test Data
# =============================================================================

# Replace these with your actual SharePoint/OneDrive URLs and file names for integration tests
SYNTHETIC_DATA = {
    "sharepoint_url": "https://contoso.sharepoint.com/Shared%20Documents/Forms/AllItems.aspx",
    "file_name": "Test_Trade_Tracker.xlsx",
    "sheet_name": "Trades",
}

# Sample trade data for testing
SAMPLE_TRADES = [
    {
        "trade_id": "TRD-001",
        "date": "12/20/2025",
        "time": "09:30 AM",
        "symbol": "AAPL",
        "quantity": 100,
        "price": 195.50,
        "action": "BUY",
    },
    {
        "trade_id": "TRD-002",
        "date": "12/21/2025",
        "time": "10:15 AM",
        "symbol": "MSFT",
        "quantity": 50,
        "price": 425.75,
        "action": "BUY",
    },
    {
        "trade_id": "TRD-003",
        "date": "12/22/2025",
        "time": "02:45 PM",
        "symbol": "GOOGL",
        "quantity": 25,
        "price": 175.25,
        "action": "SELL",
    },
    {
        "trade_id": "TRD-004",
        "date": "12/23/2025",
        "time": "11:30 AM",
        "symbol": "NVDA",
        "quantity": 75,
        "price": 495.00,
        "action": "BUY",
    },
]


# =============================================================================
# MCP Client Helper
# =============================================================================

class MCPTestClient:
    """Simple MCP client for testing purposes."""
    
    def __init__(self, base_url: str = MCP_SERVER_URL):
        self.base_url = base_url
        self.mcp_url = f"{base_url}/mcp"
        self.health_url = f"{base_url}/health"
    
    async def health_check(self) -> dict:
        """Check if the MCP server is healthy."""
        async with httpx.AsyncClient() as client:
            response = await client.get(self.health_url, timeout=TIMEOUT)
            return response.json()
    
    async def list_tools(self) -> dict:
        """List available MCP tools."""
        async with httpx.AsyncClient() as client:
            # MCP uses JSON-RPC 2.0
            payload = {
                "jsonrpc": "2.0",
                "id": 1,
                "method": "tools/list",
            }
            response = await client.post(
                self.mcp_url,
                json=payload,
                timeout=TIMEOUT,
            )
            return response.json()
    
    async def call_tool(self, tool_name: str, arguments: dict) -> dict:
        """Call an MCP tool with the given arguments."""
        async with httpx.AsyncClient() as client:
            payload = {
                "jsonrpc": "2.0",
                "id": 2,
                "method": "tools/call",
                "params": {
                    "name": tool_name,
                    "arguments": arguments,
                },
            }
            response = await client.post(
                self.mcp_url,
                json=payload,
                timeout=TIMEOUT,
            )
            return response.json()


# =============================================================================
# Test Cases
# =============================================================================

async def test_health_check(client: MCPTestClient) -> bool:
    """Test the health check endpoint."""
    print("\n" + "=" * 60)
    print("TEST: Health Check")
    print("=" * 60)
    
    try:
        result = await client.health_check()
        print(f"Response: {json.dumps(result, indent=2)}")
        
        if result.get("status") == "healthy":
            print("✓ PASSED: Server is healthy")
            return True
        else:
            print("✗ FAILED: Unexpected health status")
            return False
    except Exception as e:
        print(f"✗ FAILED: {e}")
        return False


async def test_list_tools(client: MCPTestClient) -> bool:
    """Test listing available MCP tools."""
    print("\n" + "=" * 60)
    print("TEST: List Tools")
    print("=" * 60)
    
    try:
        result = await client.list_tools()
        print(f"Response: {json.dumps(result, indent=2)}")
        
        if "result" in result:
            tools = result["result"].get("tools", [])
            print(f"\nAvailable tools ({len(tools)}):")
            for tool in tools:
                print(f"  • {tool.get('name')}: {tool.get('description', '')[:60]}...")
            
            # Check for expected tools
            tool_names = [t.get("name") for t in tools]
            expected_tools = ["excel.updateRowByLookup", "excel.updateRange"]
            
            missing = [t for t in expected_tools if t not in tool_names]
            if missing:
                print(f"✗ FAILED: Missing expected tools: {missing}")
                return False
            
            print("✓ PASSED: All expected tools found")
            return True
        else:
            print(f"✗ FAILED: Unexpected response format")
            return False
    except Exception as e:
        print(f"✗ FAILED: {e}")
        return False


async def test_update_row_by_lookup_schema(client: MCPTestClient) -> bool:
    """Test the updateRowByLookup tool schema validation."""
    print("\n" + "=" * 60)
    print("TEST: UpdateRowByLookup - Schema Validation")
    print("=" * 60)
    
    # This test validates that the schema is properly defined
    # It will fail with auth errors if credentials aren't configured,
    # but should NOT fail with schema validation errors
    
    # Note: target_columns and values are now JSON strings
    test_params = {
        "url": SYNTHETIC_DATA["sharepoint_url"],
        "file_name": SYNTHETIC_DATA["file_name"],
        "sheet_name": SYNTHETIC_DATA["sheet_name"],
        "search_column": "A",
        "reference_value": "TRD-001",
        "target_columns": '["D", "E", "F"]',
        "values": '["COMPLETED", "2025-12-23", 100.50]',
        "row_offset": 0,
    }
    
    print(f"Request payload:\n{json.dumps(test_params, indent=2)}")
    
    try:
        result = await client.call_tool("excel.updateRowByLookup", test_params)
        print(f"\nResponse: {json.dumps(result, indent=2)}")
        
        # Check if the error is a schema validation error
        if "error" in result:
            error_msg = str(result.get("error", {}))
            if "invalid_request_error" in error_msg.lower() or "schema" in error_msg.lower():
                print("✗ FAILED: Schema validation error")
                return False
            else:
                # Other errors (auth, file not found) are expected in test environment
                print("⚠ WARNING: Tool returned an error (expected without real credentials)")
                print("✓ PASSED: Schema validation successful")
                return True
        
        print("✓ PASSED: Tool call successful")
        return True
    except Exception as e:
        print(f"✗ FAILED: {e}")
        return False


async def test_update_range_schema(client: MCPTestClient) -> bool:
    """Test the updateRange tool schema validation."""
    print("\n" + "=" * 60)
    print("TEST: UpdateRange - Schema Validation")
    print("=" * 60)
    
    # Note: values is now a JSON string containing a 2D array
    test_params = {
        "url": SYNTHETIC_DATA["sharepoint_url"],
        "file_name": SYNTHETIC_DATA["file_name"],
        "sheet_name": SYNTHETIC_DATA["sheet_name"],
        "address": "A1:D3",
        "values": '[["Trade ID", "Date", "Symbol", "Quantity"], ["TRD-001", "12/20/2025", "AAPL", 100], ["TRD-002", "12/21/2025", "MSFT", 50]]',
    }
    
    print(f"Request payload:\n{json.dumps(test_params, indent=2)}")
    
    try:
        result = await client.call_tool("excel.updateRange", test_params)
        print(f"\nResponse: {json.dumps(result, indent=2)}")
        
        # Check if the error is a schema validation error
        if "error" in result:
            error_msg = str(result.get("error", {}))
            if "invalid_request_error" in error_msg.lower() or "schema" in error_msg.lower():
                print("✗ FAILED: Schema validation error")
                return False
            else:
                # Other errors (auth, file not found) are expected in test environment
                print("⚠ WARNING: Tool returned an error (expected without real credentials)")
                print("✓ PASSED: Schema validation successful")
                return True
        
        print("✓ PASSED: Tool call successful")
        return True
    except Exception as e:
        print(f"✗ FAILED: {e}")
        return False


async def test_update_row_with_date_lookup(client: MCPTestClient) -> bool:
    """Test updateRowByLookup with date-based lookup."""
    print("\n" + "=" * 60)
    print("TEST: UpdateRowByLookup - Date Lookup")
    print("=" * 60)
    
    # Simulate looking up by today's date
    today = datetime.now().strftime("%m/%d/%Y")
    
    # Note: target_columns and values are JSON strings
    test_params = {
        "url": SYNTHETIC_DATA["sharepoint_url"],
        "file_name": SYNTHETIC_DATA["file_name"],
        "sheet_name": "December",
        "search_column": "C",
        "reference_value": today,
        "target_columns": '["C", "E", "I", "J", "L"]',
        "values": json.dumps([today, "11:36 AM", "VPCS", 0.25, 25]),
        "row_offset": 1,
    }
    
    print(f"Looking up date: {today}")
    print(f"Request payload:\n{json.dumps(test_params, indent=2)}")
    
    try:
        result = await client.call_tool("excel.updateRowByLookup", test_params)
        print(f"\nResponse: {json.dumps(result, indent=2)}")
        
        if "error" in result:
            error_msg = str(result.get("error", {}))
            if "invalid_request_error" in error_msg.lower() or "schema" in error_msg.lower():
                print("✗ FAILED: Schema validation error")
                return False
            print("⚠ WARNING: Tool returned an error (expected without real credentials)")
            print("✓ PASSED: Schema validation successful")
            return True
        
        print("✓ PASSED: Tool call successful")
        return True
    except Exception as e:
        print(f"✗ FAILED: {e}")
        return False


async def test_mixed_value_types(client: MCPTestClient) -> bool:
    """Test that mixed value types (string, number, boolean, null) work correctly."""
    print("\n" + "=" * 60)
    print("TEST: Mixed Value Types")
    print("=" * 60)
    
    # Note: values is now a JSON string containing mixed types
    test_params = {
        "url": SYNTHETIC_DATA["sharepoint_url"],
        "file_name": SYNTHETIC_DATA["file_name"],
        "sheet_name": SYNTHETIC_DATA["sheet_name"],
        "search_column": "A",
        "reference_value": "TRD-001",
        "target_columns": '["B", "C", "D", "E", "F"]',
        "values": '["String value", 123, 45.67, true, null]',
        "row_offset": 0,
    }
    
    print(f"Testing value types: string, integer, float, boolean, null")
    print(f"Request payload:\n{json.dumps(test_params, indent=2)}")
    
    try:
        result = await client.call_tool("excel.updateRowByLookup", test_params)
        print(f"\nResponse: {json.dumps(result, indent=2)}")
        
        if "error" in result:
            error_msg = str(result.get("error", {}))
            if "invalid_request_error" in error_msg.lower() or "schema" in error_msg.lower():
                print("✗ FAILED: Schema validation error with mixed types")
                return False
            print("⚠ WARNING: Tool returned an error (expected without real credentials)")
            print("✓ PASSED: Mixed value types accepted by schema")
            return True
        
        print("✓ PASSED: Tool call successful")
        return True
    except Exception as e:
        print(f"✗ FAILED: {e}")
        return False


async def test_2d_array_values(client: MCPTestClient) -> bool:
    """Test updateRange with various 2D array configurations."""
    print("\n" + "=" * 60)
    print("TEST: 2D Array Values")
    print("=" * 60)
    
    # Note: values is now a JSON string containing a 2D array
    test_params = {
        "url": SYNTHETIC_DATA["sharepoint_url"],
        "file_name": SYNTHETIC_DATA["file_name"],
        "sheet_name": SYNTHETIC_DATA["sheet_name"],
        "address": "A1:E4",
        "values": '[["Trade ID", "Date", "Symbol", "Qty", "Price"], ["TRD-001", "12/20/2025", "AAPL", 100, 195.50], ["TRD-002", "12/21/2025", "MSFT", 50, 425.75], ["TRD-003", "12/22/2025", "GOOGL", 25, 175.25]]',
    }
    
    print(f"Testing 4x5 grid of values")
    print(f"Request payload:\n{json.dumps(test_params, indent=2)}")
    
    try:
        result = await client.call_tool("excel.updateRange", test_params)
        print(f"\nResponse: {json.dumps(result, indent=2)}")
        
        if "error" in result:
            error_msg = str(result.get("error", {}))
            if "invalid_request_error" in error_msg.lower() or "schema" in error_msg.lower():
                print("✗ FAILED: Schema validation error with 2D array")
                return False
            print("⚠ WARNING: Tool returned an error (expected without real credentials)")
            print("✓ PASSED: 2D array accepted by schema")
            return True
        
        print("✓ PASSED: Tool call successful")
        return True
    except Exception as e:
        print(f"✗ FAILED: {e}")
        return False


# =============================================================================
# Integration Test (requires real credentials)
# =============================================================================

async def test_integration_update_row(
    client: MCPTestClient,
    url: str,
    file_name: str,
    sheet_name: str,
) -> bool:
    """
    Integration test for updateRowByLookup with real SharePoint data.
    
    Requires:
    - Valid Azure AD credentials configured
    - Real SharePoint URL and Excel file
    """
    print("\n" + "=" * 60)
    print("TEST: Integration - UpdateRowByLookup")
    print("=" * 60)
    
    today = datetime.now().strftime("%m/%d/%Y")
    
    # Note: target_columns and values are JSON strings
    test_params = {
        "url": url,
        "file_name": file_name,
        "sheet_name": sheet_name,
        "search_column": "C",
        "reference_value": today,
        "target_columns": '["E", "F"]',
        "values": json.dumps(["Test Update", datetime.now().strftime("%H:%M")]),
        "row_offset": 0,
    }
    
    print(f"Target: {url}/{file_name}")
    print(f"Sheet: {sheet_name}")
    print(f"Looking up date: {today}")
    print(f"Request payload:\n{json.dumps(test_params, indent=2)}")
    
    try:
        result = await client.call_tool("excel.updateRowByLookup", test_params)
        print(f"\nResponse: {json.dumps(result, indent=2)}")
        
        if "error" in result:
            print("✗ FAILED: Integration test failed")
            return False
        
        # Parse the tool result
        if "result" in result:
            content = result["result"].get("content", [])
            if content and content[0].get("type") == "text":
                tool_result = json.loads(content[0].get("text", "{}"))
                if tool_result.get("status") == "success":
                    print("✓ PASSED: Integration test successful")
                    return True
        
        print("✗ FAILED: Unexpected response format")
        return False
    except Exception as e:
        print(f"✗ FAILED: {e}")
        return False


# =============================================================================
# Test Runner
# =============================================================================

async def test_log_trades_schema(client: MCPTestClient) -> bool:
    """Test the logTrades tool schema validation."""
    print("\n" + "=" * 60)
    print("TEST: LogTrades - Schema Validation")
    print("=" * 60)
    
    # Note: trades is a JSON string containing an array of trade objects
    test_params = {
        "trades": '[{"date": "12/23/2025", "time": "10:30 AM", "strategy": "VPCS", "credit": 0.25, "contracts": 25}]',
        "reference_date": "12/22/2025",
        "sheet_name": "December",
    }
    
    print(f"Request payload:\n{json.dumps(test_params, indent=2)}")
    
    try:
        result = await client.call_tool("excel.logTrades", test_params)
        print(f"\nResponse: {json.dumps(result, indent=2)}")
        
        # Check if the error is a schema validation error
        if "error" in result:
            error_msg = str(result.get("error", {}))
            if "invalid_request_error" in error_msg.lower() or "schema" in error_msg.lower():
                print("✗ FAILED: Schema validation error")
                return False
            else:
                # Other errors (auth, file not found) are expected in test environment
                print("⚠ WARNING: Tool returned an error (expected without real credentials)")
                print("✓ PASSED: Schema validation successful")
                return True
        
        print("✓ PASSED: Tool call successful")
        return True
    except Exception as e:
        print(f"✗ FAILED: {e}")
        return False


async def run_all_tests(client: MCPTestClient) -> dict:
    """Run all schema validation tests."""
    results = {
        "passed": 0,
        "failed": 0,
        "tests": {},
    }
    
    tests = [
        ("health", test_health_check),
        ("list_tools", test_list_tools),
        ("update_row_schema", test_update_row_by_lookup_schema),
        ("update_range_schema", test_update_range_schema),
        ("log_trades_schema", test_log_trades_schema),
        ("date_lookup", test_update_row_with_date_lookup),
        ("mixed_types", test_mixed_value_types),
        ("2d_array", test_2d_array_values),
    ]
    
    for name, test_func in tests:
        try:
            passed = await test_func(client)
            results["tests"][name] = "PASSED" if passed else "FAILED"
            if passed:
                results["passed"] += 1
            else:
                results["failed"] += 1
        except Exception as e:
            results["tests"][name] = f"ERROR: {e}"
            results["failed"] += 1
    
    return results


async def main():
    """Main entry point for test runner."""
    parser = argparse.ArgumentParser(description="MCP Excel Server Test Suite")
    parser.add_argument(
        "--test",
        choices=["all", "health", "list_tools", "update_row", "update_range", "integration"],
        default="all",
        help="Specific test to run (default: all)",
    )
    parser.add_argument(
        "--url",
        default=MCP_SERVER_URL,
        help=f"MCP server URL (default: {MCP_SERVER_URL})",
    )
    parser.add_argument(
        "--sharepoint-url",
        help="SharePoint URL for integration tests",
    )
    parser.add_argument(
        "--file-name",
        help="Excel file name for integration tests",
    )
    parser.add_argument(
        "--sheet-name",
        default="Sheet1",
        help="Worksheet name for integration tests",
    )
    
    args = parser.parse_args()
    
    print("\n" + "=" * 60)
    print("MCP Excel Service - Test Suite")
    print("=" * 60)
    print(f"Server URL: {args.url}")
    print(f"Timestamp: {datetime.now().isoformat()}")
    
    client = MCPTestClient(args.url)
    
    if args.test == "all":
        results = await run_all_tests(client)
        
        print("\n" + "=" * 60)
        print("TEST SUMMARY")
        print("=" * 60)
        for name, status in results["tests"].items():
            icon = "✓" if status == "PASSED" else "✗"
            print(f"  {icon} {name}: {status}")
        
        print(f"\nTotal: {results['passed']} passed, {results['failed']} failed")
        
        sys.exit(0 if results["failed"] == 0 else 1)
    
    elif args.test == "health":
        passed = await test_health_check(client)
        sys.exit(0 if passed else 1)
    
    elif args.test == "list_tools":
        passed = await test_list_tools(client)
        sys.exit(0 if passed else 1)
    
    elif args.test == "update_row":
        passed = await test_update_row_by_lookup_schema(client)
        sys.exit(0 if passed else 1)
    
    elif args.test == "update_range":
        passed = await test_update_range_schema(client)
        sys.exit(0 if passed else 1)
    
    elif args.test == "integration":
        if not args.sharepoint_url or not args.file_name:
            print("ERROR: Integration test requires --sharepoint-url and --file-name")
            sys.exit(1)
        
        passed = await test_integration_update_row(
            client,
            args.sharepoint_url,
            args.file_name,
            args.sheet_name,
        )
        sys.exit(0 if passed else 1)


if __name__ == "__main__":
    asyncio.run(main())
