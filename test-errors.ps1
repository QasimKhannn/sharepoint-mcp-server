# Test script for error handling

Write-Host "====== SharePoint MCP Error Handling Tests ======`n" -ForegroundColor Cyan

# Test 1: Test tools/list (discovery mode - no auth required)
Write-Host "Test 1: Tools Discovery (no auth required)" -ForegroundColor Yellow
$json1 = @{
    jsonrpc = "2.0"
    id = 1
    method = "tools/list"
} | ConvertTo-Json -Depth 10

$response1 = Invoke-RestMethod -Uri "http://localhost:8080/mcp" `
  -Method POST `
  -ContentType "application/json" `
  -Body $json1 -ErrorAction SilentlyContinue

Write-Host "Response: $(($response1.result.tools | Measure-Object).Count) tools available"
Write-Host ""

# Test 2: Tool call without token
Write-Host "Test 2: Missing Authentication Token" -ForegroundColor Yellow
$json2 = @{
    jsonrpc = "2.0"
    id = 2
    method = "tools/call"
    params = @{
        name = "get_current_user"
        arguments = @{}
    }
} | ConvertTo-Json -Depth 10

$response2 = Invoke-RestMethod -Uri "http://localhost:8080/mcp" `
  -Method POST `
  -ContentType "application/json" `
  -Body $json2 -ErrorAction SilentlyContinue

Write-Host "Response:"
Write-Host ($response2 | ConvertTo-Json -Depth 5)
Write-Host ""

# Test 3: Missing required parameter
Write-Host "Test 3: Missing Required Parameter (driveId)" -ForegroundColor Yellow
$json3 = @{
    jsonrpc = "2.0"
    id = 3
    method = "tools/call"
    params = @{
        name = "read_file_content"
        arguments = @{
            itemId = "test-item"
        }
    }
} | ConvertTo-Json -Depth 10

$response3 = Invoke-RestMethod -Uri "http://localhost:8080/mcp" `
  -Method POST `
  -Headers @{ "Authorization" = "Bearer test-token" } `
  -ContentType "application/json" `
  -Body $json3 -ErrorAction SilentlyContinue

Write-Host "Response:"
Write-Host ($response3 | ConvertTo-Json -Depth 5)
Write-Host ""

# Test 4: Unsupported file type (simulated)
Write-Host "Test 4: Simulated Unsupported File Type Error" -ForegroundColor Yellow
Write-Host "When attempting to read a .exe or .zip file, the response would be:"
Write-Host @'
{
  "jsonrpc": "2.0",
  "id": 4,
  "error": {
    "code": 400,
    "message": "Unsupported file type: .exe. Supported formats: DOCX, PDF, XLSX, XLS, PPTX, TXT, MD, CSV, JSON"
  }
}
'@
Write-Host ""

# Test 5: Invalid parameter value
Write-Host "Test 5: Invalid Sharing Link Type Parameter" -ForegroundColor Yellow
$json5 = @{
    jsonrpc = "2.0"
    id = 5
    method = "tools/call"
    params = @{
        name = "create_sharing_link"
        arguments = @{
            driveId = "test-drive"
            itemId = "test-item"
            type = "invalid-type"
        }
    }
} | ConvertTo-Json -Depth 10

$response5 = Invoke-RestMethod -Uri "http://localhost:8080/mcp" `
  -Method POST `
  -Headers @{ "Authorization" = "Bearer test-token" } `
  -ContentType "application/json" `
  -Body $json5 -ErrorAction SilentlyContinue

Write-Host "Response:"
Write-Host ($response5 | ConvertTo-Json -Depth 5)
Write-Host ""

Write-Host "====== All Error Handling Tests Complete ======" -ForegroundColor Cyan
