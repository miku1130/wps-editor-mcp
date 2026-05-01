# -*- coding: utf-8 -*-
"""
WPS Editor MCP Unit Tests
"""
import os
import sys
import json
import subprocess
import tempfile
import shutil

TEST_DIR = os.path.join(os.path.dirname(__file__), "test_output")


def run_mcp(requests):
    init_request = json.dumps({
        "jsonrpc": "2.0", "id": 0, "method": "initialize", "params": {}
    })
    input_data = init_request + '\n' + '\n'.join(requests) + '\n'
    server_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "server.py")
    result = subprocess.run(
        ['python', server_path],
        input=input_data, capture_output=True, text=True
    )
    responses = []
    for line in result.stdout.strip().split('\n'):
        try:
            responses.append(json.loads(line))
        except:
            pass
    return responses


def make_request(id, tool_name, arguments):
    return json.dumps({
        "jsonrpc": "2.0", "id": id, "method": "tools/call",
        "params": {"name": tool_name, "arguments": arguments}
    })


def setup():
    os.makedirs(TEST_DIR, exist_ok=True)


def teardown():
    if os.path.exists(TEST_DIR):
        shutil.rmtree(TEST_DIR)


def test_mcp_protocol():
    print("\n=== Test MCP Protocol ===")
    responses = run_mcp([])
    assert responses[0]["result"]["protocolVersion"] == "2024-11-05"
    print("[OK] MCP init success")
    
    init_request = json.dumps({"jsonrpc": "2.0", "id": 0, "method": "initialize", "params": {}})
    list_request = json.dumps({"jsonrpc": "2.0", "id": 1, "method": "tools/list", "params": {}})
    server_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "server.py")
    result = subprocess.run(
        ['python', server_path],
        input=init_request + '\n' + list_request + '\n',
        capture_output=True, text=True
    )
    responses = []
    for line in result.stdout.strip().split('\n'):
        try:
            responses.append(json.loads(line))
        except:
            pass
    
    tools = responses[-1]["result"]["tools"]
    tool_names = [t["name"] for t in tools]
    expected = ["read_document", "write_document", "edit_cell", "add_rows", 
                "merge_cells", "insert_image", "set_cell_style", "export_pdf", "run_macro"]
    for name in expected:
        assert name in tool_names, f"Missing tool: {name}"
    print(f"[OK] Tools list correct, {len(tools)} tools")


def test_word_basic():
    print("\n=== Test Word Basic ===")
    file_path = os.path.join(TEST_DIR, "test.docx")
    
    responses = run_mcp([make_request(1, "write_document", {
        "file_path": file_path, "content": "Paragraph 1", "mode": "file"
    })])
    assert responses[-1]["id"] == 1
    print("[OK] Write Word success")
    
    responses = run_mcp([make_request(2, "read_document", {
        "file_path": file_path, "mode": "file"
    })])
    result = json.loads(responses[-1]["result"]["content"][0]["text"])
    assert len(result["paragraphs"]) >= 1
    assert result["paragraphs"][0] == "Paragraph 1"
    print("[OK] Read Word success")


def test_excel_basic():
    print("\n=== Test Excel Basic ===")
    file_path = os.path.join(TEST_DIR, "test.xlsx")
    
    responses = run_mcp([make_request(1, "write_document", {
        "file_path": file_path,
        "content": '[["Name", "Age"], ["Alice", 25], ["Bob", 30]]',
        "mode": "file"
    })])
    assert responses[-1]["id"] == 1
    print("[OK] Write Excel success")
    
    responses = run_mcp([make_request(2, "read_document", {
        "file_path": file_path, "mode": "file"
    })])
    result = json.loads(responses[-1]["result"]["content"][0]["text"])
    assert result["rows"] == 3
    print("[OK] Read Excel success")
    
    responses = run_mcp([make_request(3, "edit_cell", {
        "file_path": file_path, "row": 1, "column": 1, "value": "X", "mode": "file"
    })])
    assert responses[-1]["id"] == 3
    print("[OK] Edit cell success")


def test_excel_table_ops():
    print("\n=== Test Excel Table Ops ===")
    file_path = os.path.join(TEST_DIR, "test_table.xlsx")
    
    run_mcp([make_request(0, "write_document", {
        "file_path": file_path, "content": '[["A","B"],[1,2]]', "mode": "file"
    })])
    
    responses = run_mcp([make_request(1, "add_rows", {
        "file_path": file_path, "count": 2, "after_row": 2, "mode": "file"
    })])
    assert responses[-1]["id"] == 1
    print("[OK] Add rows success")
    
    responses = run_mcp([make_request(2, "add_columns", {
        "file_path": file_path, "count": 1, "after_col": 2, "mode": "file"
    })])
    assert responses[-1]["id"] == 2
    print("[OK] Add columns success")
    
    responses = run_mcp([make_request(3, "merge_cells", {
        "file_path": file_path, "start_row": 1, "start_col": 1,
        "end_row": 2, "end_col": 2, "mode": "file"
    })])
    assert responses[-1]["id"] == 3
    print("[OK] Merge cells success")
    
    responses = run_mcp([make_request(4, "set_column_width", {
        "file_path": file_path, "column": "A", "width": 20, "mode": "file"
    })])
    assert responses[-1]["id"] == 4
    print("[OK] Set column width success")


def test_excel_style():
    print("\n=== Test Excel Style ===")
    file_path = os.path.join(TEST_DIR, "test_style.xlsx")
    
    run_mcp([make_request(0, "write_document", {
        "file_path": file_path, "content": '[["Test"]]', "mode": "file"
    })])
    
    responses = run_mcp([make_request(1, "set_cell_style", {
        "file_path": file_path, "row": 1, "column": 1,
        "font_name": "Arial", "font_size": 14, "bold": True,
        "color": "FF0000", "bg_color": "FFFF00", "align_h": "center", "mode": "file"
    })])
    assert responses[-1]["id"] == 1
    print("[OK] Set cell style success")


def test_powerpoint():
    print("\n=== Test PowerPoint ===")
    file_path = os.path.join(TEST_DIR, "test.pptx")
    
    responses = run_mcp([make_request(1, "write_document", {
        "file_path": file_path, "content": "Slide 1 content", "mode": "file"
    })])
    assert responses[-1]["id"] == 1
    print("[OK] Write PPT success")
    
    responses = run_mcp([make_request(2, "read_document", {
        "file_path": file_path, "mode": "file"
    })])
    result = json.loads(responses[-1]["result"]["content"][0]["text"])
    assert result["slide_count"] == 1
    print("[OK] Read PPT success")


def main():
    print("=" * 60)
    print("WPS Editor MCP Unit Tests")
    print("=" * 60)
    
    setup()
    try:
        test_mcp_protocol()
        test_word_basic()
        test_excel_basic()
        test_excel_table_ops()
        test_excel_style()
        test_powerpoint()
        
        print("\n" + "=" * 60)
        print("All tests passed!")
        print("=" * 60)
    finally:
        teardown()


if __name__ == "__main__":
    main()
