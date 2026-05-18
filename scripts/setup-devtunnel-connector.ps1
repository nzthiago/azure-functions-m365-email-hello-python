<#
.SYNOPSIS
  Stand up a Dev Tunnel + Connector Namespace + Office 365 connection +
  trigger config that calls back into a local `func start` host for the
  hello-microsoft365email-connector sample.

.DESCRIPTION
  Mirrors scripts/setup-devtunnel-connector.sh for Windows / PowerShell.

  Steps:
    1. Create (or reuse) an anonymous Dev Tunnel on port 7071 and host it.
    2. Create a Connector Namespace in westcentralus (if missing).
    3. Create an Office 365 connection and prompt for OAuth consent.
    4. List trigger operations available on the connection.
    5. Create a trigger-config whose callbackUrl points at the devtunnel URL
       + /runtime/webhooks/connector?functionName=<FunctionName>.

.EXAMPLE
  ./scripts/setup-devtunnel-connector.ps1 `
    -ResourceGroup  hello-m365-rg `
    -Namespace      hello-m365-ns `
    -Connection     office365-connection `
    -TriggerConfig  on-new-email `
    -FunctionName   OnNewEmail
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $ResourceGroup,
    [Parameter(Mandatory = $true)] [string] $Namespace,
    [string] $Connection    = "office365-connection",
    [string] $TriggerConfig = "on-new-email",
    [string] $FunctionName  = "OnNewEmail",
    [string] $Location      = "westcentralus",
    [int]    $Port          = 7071,
    [string] $TunnelId      = ""
)

$ErrorActionPreference = "Stop"

function Log  ($m) { Write-Host ">>> $m" -ForegroundColor Cyan }
function Ok   ($m) { Write-Host "[OK] $m"  -ForegroundColor Green }
function Warn ($m) { Write-Host "[!]  $m"  -ForegroundColor Yellow }

foreach ($cmd in @("az", "devtunnel")) {
    if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) {
        throw "$cmd CLI not found in PATH."
    }
}

# --- 1. Dev tunnel ----------------------------------------------------------
Log "Ensuring devtunnel login..."
try { devtunnel user show *> $null } catch { devtunnel user login }

if (-not $TunnelId) {
    Log "Creating anonymous devtunnel on port $Port..."
    $createJson = devtunnel create --allow-anonymous --json | ConvertFrom-Json
    $TunnelId = $createJson.tunnel.tunnelId
    devtunnel port create $TunnelId -p $Port --protocol http | Out-Null
    Ok "Created tunnel: $TunnelId"
} else {
    Log "Reusing existing tunnel: $TunnelId"
}

# `devtunnel show` only populates the public port URI while the tunnel is
# being hosted, so start `devtunnel host` in the background and poll until the
# URL appears.
$HostLog = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "devtunnel-host-$([System.Guid]::NewGuid().ToString('N')).log")
Log "Starting 'devtunnel host $TunnelId' in background (log: $HostLog)..."
$HostProc = Start-Process -FilePath "devtunnel" `
    -ArgumentList @("host", $TunnelId) `
    -RedirectStandardOutput $HostLog `
    -RedirectStandardError  "$HostLog.err" `
    -NoNewWindow -PassThru

$TunnelUri = $null
for ($i = 0; $i -lt 30; $i++) {
    Start-Sleep -Seconds 1
    try {
        $showJson  = devtunnel show $TunnelId --json 2>$null | ConvertFrom-Json
        $portEntry = $showJson.tunnel.ports | Where-Object { $_.portNumber -eq $Port }
        if ($portEntry -and $portEntry.portUri) {
            $TunnelUri = ($portEntry.portUri).TrimEnd("/")
            break
        }
    } catch { }
}
if (-not $TunnelUri) {
    if ($HostProc -and -not $HostProc.HasExited) { Stop-Process -Id $HostProc.Id -Force }
    throw "Timed out waiting for devtunnel public URL. Host log: $HostLog"
}
Ok "Public tunnel URL: $TunnelUri"

$CallbackUrl = "$TunnelUri/runtime/webhooks/connector?functionName=$FunctionName"
Log "Callback URL: $CallbackUrl"

# --- 2. Resource group + namespace -----------------------------------------
Log "Ensuring resource group $ResourceGroup ($Location)..."
az group create -n $ResourceGroup -l $Location -o none

Log "Ensuring connector namespace $Namespace..."
$nsExists = $true
try { az connector-namespace show -n $Namespace -g $ResourceGroup -o none 2>$null } catch { $nsExists = $false }
if (-not $nsExists) {
    az connector-namespace create `
        --name $Namespace `
        --resource-group $ResourceGroup `
        --location $Location -o none
    Ok "Created namespace $Namespace"
} else {
    Ok "Namespace already exists"
}

# --- 3. Connection + OAuth consent -----------------------------------------
Log "Ensuring Office 365 connection $Connection..."
$connExists = $true
try {
    az connector-namespace connection show `
        --namespace-name $Namespace -g $ResourceGroup -n $Connection -o none 2>$null
} catch { $connExists = $false }
if (-not $connExists) {
    az connector-namespace connection create `
        --namespace-name $Namespace `
        --resource-group $ResourceGroup `
        --name $Connection `
        --available-connector office365 -o none
    Ok "Created connection $Connection"
}

Warn "Authorizing connection - a browser window will open for Office 365 OAuth consent..."
az connector-namespace connection authorize `
    --namespace-name $Namespace `
    --resource-group $ResourceGroup `
    --name $Connection

# --- 4. Verify trigger operation is discoverable ---------------------------
Log "Listing trigger operations on the connection..."
az connector-namespace connection operation list `
    --namespace-name $Namespace `
    --resource-group $ResourceGroup `
    -n $Connection `
    --operation-type trigger -o table

# --- 5. Trigger config ------------------------------------------------------
Log "Creating trigger config $TriggerConfig -> $CallbackUrl..."
$tcExists = $true
try {
    az connector-namespace trigger-config show `
        --namespace-name $Namespace -g $ResourceGroup -n $TriggerConfig -o none 2>$null
} catch { $tcExists = $false }
if ($tcExists) {
    Warn "Trigger config already exists - deleting and recreating to update callbackUrl"
    az connector-namespace trigger-config delete `
        --namespace-name $Namespace -g $ResourceGroup -n $TriggerConfig --yes -o none
}

az connector-namespace trigger-config create `
    --namespace-name $Namespace `
    --resource-group $ResourceGroup `
    --name $TriggerConfig `
    --available-connector office365 `
    --connection-name $Connection `
    --operation-name OnNewEmailV3 `
    --callback-url $CallbackUrl `
    --parameter folderPath=Inbox `
    -o none

Ok "Trigger config $TriggerConfig created."

Write-Host @"

----------------------------------------------------------------------
  Devtunnel is hosting in the background (PID $($HostProc.Id),
  log $HostLog). Keep this script running - Ctrl+C stops the tunnel.

  Next steps:
    1. In another terminal, start the function host:
          func start

    2. Send yourself an email at the account you just authorized.
       Watch the func start logs - you should see "Received Microsoft 365
       OnNewEmail trigger" followed by the email subject and from.

    3. Inspect recent gateway -> function deliveries:
          az connector-namespace trigger-config run list ``
            -g $ResourceGroup --namespace-name $Namespace ``
            --trigger-config-name $TriggerConfig -o table

  Tunnel ID:   $TunnelId
  Tunnel URL:  $TunnelUri
  Callback:    $CallbackUrl
----------------------------------------------------------------------
"@ -ForegroundColor Yellow

Log "Tailing devtunnel host (Ctrl+C to stop)..."
try {
    Wait-Process -Id $HostProc.Id
} finally {
    if ($HostProc -and -not $HostProc.HasExited) {
        Stop-Process -Id $HostProc.Id -Force -ErrorAction SilentlyContinue
    }
}
