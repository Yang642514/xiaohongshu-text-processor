param(
  [int]$Port = 10065
)

$listener = New-Object System.Net.HttpListener
$prefix = "http://127.0.0.1:$Port/"
$listener.Prefixes.Add($prefix)
$listener.Start()
Write-Host "Server started at $prefix"

try {
  while ($true) {
    $context = $listener.GetContext()
    $req = $context.Request
    $localPath = $req.Url.LocalPath.TrimStart('/')
    if ([string]::IsNullOrWhiteSpace($localPath)) { $localPath = 'ui-preview.html' }
    $path = Join-Path (Get-Location) $localPath
    if (-not (Test-Path $path)) {
      $context.Response.StatusCode = 404
      $buf = [System.Text.Encoding]::UTF8.GetBytes('Not Found')
      $context.Response.OutputStream.Write($buf,0,$buf.Length)
      $context.Response.Close()
      continue
    }
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $contentType = 'text/html'
    switch -Regex ($path.ToLower()) {
      '.*\.css$' { $contentType = 'text/css'; break }
      '.*\.js$' { $contentType = 'application/javascript'; break }
      '.*\.png$' { $contentType = 'image/png'; break }
      '.*\.(jpg|jpeg)$' { $contentType = 'image/jpeg'; break }
    }
    $context.Response.ContentType = $contentType
    $context.Response.OutputStream.Write($bytes,0,$bytes.Length)
    $context.Response.Close()
  }
}
finally {
  $listener.Stop()
}