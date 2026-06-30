# Permit Management System static server — PowerShell HttpListener (no Node required)
param([int]$Port = 8131)
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$Port/")
$listener.Start()
Write-Host "Permit System serving $root at http://localhost:$Port/"
$mime = @{ ".html"="text/html"; ".css"="text/css"; ".js"="application/javascript"; ".svg"="image/svg+xml"; ".json"="application/json"; ".png"="image/png"; ".jpg"="image/jpeg"; ".ico"="image/x-icon" }
while ($listener.IsListening) {
  $ctx = $listener.GetContext()
  $path = $ctx.Request.Url.LocalPath.TrimStart("/")
  if ([string]::IsNullOrEmpty($path)) { $path = "index.html" }
  $file = Join-Path $root $path
  if (Test-Path $file -PathType Leaf) {
    $ext = [System.IO.Path]::GetExtension($file).ToLower()
    $ct = $mime[$ext]; if (-not $ct) { $ct = "application/octet-stream" }
    $bytes = [System.IO.File]::ReadAllBytes($file)
    $ctx.Response.ContentType = $ct
    $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
  } else {
    $ctx.Response.StatusCode = 404
  }
  $ctx.Response.Close()
}
