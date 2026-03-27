param(
  [string]$InputsJson = "",
  [string]$AccessToken = ""
)

$envPath = Join-Path $PSScriptRoot ".env"
if (Test-Path $envPath) {
  Write-Host "Loading .env from $envPath"
  Get-Content $envPath | ForEach-Object {
    if ($_ -match '^(?<k>[^#=]+)=(?<v>.*)$') {
      $k = $Matches['k'].Trim()
      $v = $Matches['v']
      [Environment]::SetEnvironmentVariable($k, $v)
    }
  }
}

$python = Join-Path $PSScriptRoot ".venv/Scripts/python.exe"
if (-not (Test-Path $python)) { $python = "python" }

if ($AccessToken -or $env:ACCESS_TOKEN) {
  # Token-only mode
  if ($InputsJson) {
    if ($AccessToken) {
      & $python (Join-Path $PSScriptRoot "run_desktop_flow_token.py") --token $AccessToken --inputs $InputsJson
    } else {
      & $python (Join-Path $PSScriptRoot "run_desktop_flow_token.py") --inputs $InputsJson
    }
  } else {
    if ($AccessToken) {
      & $python (Join-Path $PSScriptRoot "run_desktop_flow_token.py") --token $AccessToken
    } else {
      & $python (Join-Path $PSScriptRoot "run_desktop_flow_token.py")
    }
  }
} else {
  # Client credentials mode
  if ($InputsJson) {
    & $python (Join-Path $PSScriptRoot "run_desktop_flow.py") --inputs $InputsJson
  } else {
    & $python (Join-Path $PSScriptRoot "run_desktop_flow.py")
  }
}

exit $LASTEXITCODE
