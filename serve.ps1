param(
  [int]$Port = 8787
)

$root = (Resolve-Path (Split-Path -Parent $MyInvocation.MyCommand.Path)).Path

# pick a free port if taken
function Test-PortInUse([int]$p){
  try {
    $l = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, $p)
    $l.Start(); $l.Stop(); return $false
  } catch { return $true }
}
if (Test-PortInUse $Port) { $Port = 0 }
if ($Port -eq 0) {
  for ($p=8787; $p -le 8899; $p++) { if (-not (Test-PortInUse $p)) { $Port=$p; break } }
}
if ($Port -eq 0) { throw "No free port found." }

$prefix = "http://localhost:$Port/"
$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add($prefix)
$listener.Start()

Write-Host "Serving: $root" -ForegroundColor Green
Write-Host "Open:    $prefix" -ForegroundColor Green

Start-Process $prefix | Out-Null

function Get-ContentType($path) {
  switch ([IO.Path]::GetExtension($path).ToLower()) {
    '.html' { 'text/html; charset=utf-8' }
    '.js'   { 'application/javascript; charset=utf-8' }
    '.css'  { 'text/css; charset=utf-8' }
    '.json' { 'application/json; charset=utf-8' }
    '.txt'  { 'text/plain; charset=utf-8' }
    '.png'  { 'image/png' }
    '.jpg' { 'image/jpeg' }
    '.jpeg' { 'image/jpeg' }
    '.svg'  { 'image/svg+xml' }
    default { 'application/octet-stream' }
  }
}

try {
  while ($listener.IsListening) {
    $ctx = $listener.GetContext()
    $req = $ctx.Request
    $res = $ctx.Response

    $rel = [Uri]::UnescapeDataString($req.Url.AbsolutePath.TrimStart('/'))
    if ([string]::IsNullOrWhiteSpace($rel)) { $rel = 'index.html' }

    # block path traversal
    if ($rel.Contains('..')) {
      $res.StatusCode = 400
      $bytes = [Text.Encoding]::UTF8.GetBytes('Bad request')
      $res.OutputStream.Write($bytes,0,$bytes.Length)
      $res.Close(); continue
    }

    $file = Join-Path $root $rel
    if (Test-Path $file -PathType Container) { $file = Join-Path $file 'index.html' }

    if (-not (Test-Path $file -PathType Leaf)) {
      $res.StatusCode = 404
      $bytes = [Text.Encoding]::UTF8.GetBytes('Not found')
      $res.OutputStream.Write($bytes,0,$bytes.Length)
      $res.Close(); continue
    }

    $res.StatusCode = 200
    $res.ContentType = (Get-ContentType $file)
    $bytes = [IO.File]::ReadAllBytes($file)
    $res.ContentLength64 = $bytes.Length
    $res.OutputStream.Write($bytes,0,$bytes.Length)
    $res.Close()
  }
}
finally {
  if ($listener) { $listener.Stop(); $listener.Close() }
}
