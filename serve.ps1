param(
  [int]$Port = 8000
)

Write-Host "Serving local clone at http://127.0.0.1:$Port/"
python -m http.server $Port
