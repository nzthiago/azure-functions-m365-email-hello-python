#!/usr/bin/env bash
# -----------------------------------------------------------------------------
# setup-devtunnel-connector.sh
#
# Stands up the wiring needed to develop the hello-microsoft365email-connector
# sample locally against a REAL Office 365 mailbox:
#
#   1. Create (or reuse) an anonymous Dev Tunnel on port 7071 and host it.
#   2. Create a Connector Namespace in westcentralus (if missing).
#   3. Create an Office 365 connection and prompt for OAuth consent.
#   4. Discover the OnNewEmailV3 trigger operation.
#   5. Create a trigger-config whose callbackUrl points at the devtunnel URL
#      + /runtime/webhooks/connector?functionName=<FunctionName>
#
# Usage:
#   ./scripts/setup-devtunnel-connector.sh \
#     --resource-group hello-m365-rg \
#     --namespace      hello-m365-ns \
#     --connection     office365-connection \
#     --trigger-config on-new-email \
#     --function-name  OnNewEmail \
#     [--location westcentralus] \
#     [--port 7071]
#
# Prereqs: az cli (+ connector-namespace extension), devtunnel cli.
# -----------------------------------------------------------------------------
set -euo pipefail

RESOURCE_GROUP=""
NAMESPACE=""
CONNECTION="office365-connection"
TRIGGER_CONFIG="on-new-email"
FUNCTION_NAME="OnNewEmail"
LOCATION="westcentralus"
PORT="7071"
TUNNEL_ID=""

while [[ $# -gt 0 ]]; do
  case "$1" in
    --resource-group) RESOURCE_GROUP="$2"; shift 2;;
    --namespace)      NAMESPACE="$2";      shift 2;;
    --connection)     CONNECTION="$2";     shift 2;;
    --trigger-config) TRIGGER_CONFIG="$2"; shift 2;;
    --function-name)  FUNCTION_NAME="$2";  shift 2;;
    --location)       LOCATION="$2";       shift 2;;
    --port)           PORT="$2";           shift 2;;
    --tunnel-id)      TUNNEL_ID="$2";      shift 2;;
    -h|--help)
      sed -n '2,30p' "$0"; exit 0;;
    *) echo "Unknown arg: $1" >&2; exit 1;;
  esac
done

if [[ -z "$RESOURCE_GROUP" || -z "$NAMESPACE" ]]; then
  echo "ERROR: --resource-group and --namespace are required." >&2
  exit 1
fi

command -v az        >/dev/null || { echo "az CLI not found"; exit 1; }
command -v devtunnel >/dev/null || {
  echo "devtunnel CLI not found. Install: https://learn.microsoft.com/azure/developer/dev-tunnels/get-started"
  exit 1
}

YELLOW='\033[1;33m'; GREEN='\033[0;32m'; CYAN='\033[0;36m'; NC='\033[0m'
log() { echo -e "${CYAN}>>>${NC} $*"; }
ok()  { echo -e "${GREEN}✓${NC}  $*"; }
warn(){ echo -e "${YELLOW}!${NC}  $*"; }

# --- 1. Dev tunnel -----------------------------------------------------------
log "Ensuring devtunnel login..."
devtunnel user show >/dev/null 2>&1 || devtunnel user login

if [[ -z "$TUNNEL_ID" ]]; then
  log "Creating anonymous devtunnel on port ${PORT}..."
  TUNNEL_ID=$(devtunnel create --allow-anonymous --json | jq -r '.tunnel.tunnelId')
  devtunnel port create "$TUNNEL_ID" -p "$PORT" --protocol http >/dev/null
  ok "Created tunnel: $TUNNEL_ID"
else
  log "Reusing existing tunnel: $TUNNEL_ID"
fi

# `devtunnel show` only populates `portUri` while the tunnel is being hosted,
# so we start `devtunnel host` in the background and poll until the URL appears.
HOST_LOG="$(mktemp -t devtunnel-host.XXXXXX.log)"
log "Starting 'devtunnel host ${TUNNEL_ID}' in background (log: ${HOST_LOG})..."
devtunnel host "$TUNNEL_ID" >"$HOST_LOG" 2>&1 &
HOST_PID=$!
# Make sure we kill the background host if the script aborts before the final exec.
cleanup() { [[ -n "${HOST_PID:-}" ]] && kill "$HOST_PID" 2>/dev/null || true; }
trap cleanup EXIT

TUNNEL_URI=""
for _ in $(seq 1 30); do
  TUNNEL_URI=$(devtunnel show "$TUNNEL_ID" --json 2>/dev/null \
    | jq -r ".tunnel.ports[]? | select(.portNumber==${PORT}) | .portUri // empty")
  [[ -n "$TUNNEL_URI" ]] && break
  sleep 1
done
TUNNEL_URI="${TUNNEL_URI%/}"
if [[ -z "$TUNNEL_URI" ]]; then
  echo "ERROR: timed out waiting for devtunnel public URL. Host log:" >&2
  cat "$HOST_LOG" >&2 || true
  exit 1
fi
ok "Public tunnel URL: ${TUNNEL_URI}"

CALLBACK_URL="${TUNNEL_URI}/runtime/webhooks/connector?functionName=${FUNCTION_NAME}"
log "Callback URL: ${CALLBACK_URL}"

# --- 2. Resource group + namespace ------------------------------------------
log "Ensuring resource group ${RESOURCE_GROUP} (${LOCATION})..."
az group create -n "$RESOURCE_GROUP" -l "$LOCATION" -o none

log "Ensuring connector namespace ${NAMESPACE}..."
if ! az connector-namespace show -n "$NAMESPACE" -g "$RESOURCE_GROUP" -o none 2>/dev/null; then
  az connector-namespace create \
    --name "$NAMESPACE" \
    --resource-group "$RESOURCE_GROUP" \
    --location "$LOCATION" -o none
  ok "Created namespace ${NAMESPACE}"
else
  ok "Namespace already exists"
fi

# --- 3. Connection + OAuth consent ------------------------------------------
log "Ensuring Office 365 connection ${CONNECTION}..."
if ! az connector-namespace connection show \
      --namespace-name "$NAMESPACE" -g "$RESOURCE_GROUP" -n "$CONNECTION" -o none 2>/dev/null; then
  az connector-namespace connection create \
    --namespace-name "$NAMESPACE" \
    --resource-group "$RESOURCE_GROUP" \
    --name "$CONNECTION" \
    --available-connector office365 -o none
  ok "Created connection ${CONNECTION}"
fi

warn "Authorizing connection — a browser window will open for Office 365 OAuth consent..."
az connector-namespace connection authorize \
  --namespace-name "$NAMESPACE" \
  --resource-group "$RESOURCE_GROUP" \
  --name "$CONNECTION"

# --- 4. Verify trigger operation is discoverable ----------------------------
log "Listing trigger operations on the connection..."
az connector-namespace connection operation list \
  --namespace-name "$NAMESPACE" \
  --resource-group "$RESOURCE_GROUP" \
  -n "$CONNECTION" \
  --operation-type trigger -o table

# --- 5. Trigger config ------------------------------------------------------
log "Creating trigger config ${TRIGGER_CONFIG} -> ${CALLBACK_URL}..."
if az connector-namespace trigger-config show \
      --namespace-name "$NAMESPACE" -g "$RESOURCE_GROUP" -n "$TRIGGER_CONFIG" -o none 2>/dev/null; then
  warn "Trigger config already exists — deleting and recreating to update callbackUrl"
  az connector-namespace trigger-config delete \
    --namespace-name "$NAMESPACE" -g "$RESOURCE_GROUP" -n "$TRIGGER_CONFIG" --yes -o none
fi

az connector-namespace trigger-config create \
  --namespace-name "$NAMESPACE" \
  --resource-group "$RESOURCE_GROUP" \
  --name "$TRIGGER_CONFIG" \
  --available-connector office365 \
  --connection-name "$CONNECTION" \
  --operation-name OnNewEmailV3 \
  --callback-url "$CALLBACK_URL" \
  --parameter folderPath=Inbox \
  -o none

ok "Trigger config ${TRIGGER_CONFIG} created."

cat <<EOF

────────────────────────────────────────────────────────────────────────
  Devtunnel is hosting in the background (PID ${HOST_PID}, log ${HOST_LOG}).
  Keep this script running — Ctrl+C will stop the tunnel.

  Next steps:
    1. In another terminal, start the function host:
          func start

    2. Send yourself an email at the account you just authorized.
       Watch the func start logs — you should see "Received Microsoft 365
       OnNewEmail trigger" followed by the email subject and from.

    3. Inspect recent gateway -> function deliveries:
          az connector-namespace trigger-config run list \\
            -g ${RESOURCE_GROUP} --namespace-name ${NAMESPACE} \\
            --trigger-config-name ${TRIGGER_CONFIG} -o table

  Tunnel ID:   ${TUNNEL_ID}
  Tunnel URL:  ${TUNNEL_URI}
  Callback:    ${CALLBACK_URL}
────────────────────────────────────────────────────────────────────────
EOF

log "Tailing devtunnel host (Ctrl+C to stop)..."
trap - EXIT
trap 'kill "$HOST_PID" 2>/dev/null || true; exit 0' INT TERM
wait "$HOST_PID"
