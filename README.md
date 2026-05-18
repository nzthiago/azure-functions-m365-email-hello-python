# Hello, Microsoft 365 Email Connector — Python

A minimal **"hello world"** Azure Functions app (**Python 3.13**, v2 programming
model) that prints every new email arriving in an Office 365 Outlook inbox to the
function log, using the new [`ConnectorTrigger`](https://github.com/Azure/azure-functions-connector-extension)
binding from the **Azure Functions Connectors private preview**.

> ⚠️ **Private preview** — this sample uses the `azurefunctions-extensions-connectors`
> preview package (from PyPI) and the experimental `connector-namespace`
> Azure CLI extension. See the
> [private preview instructions](https://gist.github.com/nzthiago/35799f1dc3b56d8f8915b1427e086c6c)
> for the full picture (NDA / not for production).

For a richer end-to-end sample (Teams card, sender enrichment, IaC via `azd`), see
[`FunctionAppConnectorsEmailProcessor`](https://github.com/nzthiago/FunctionAppConnectorsEmailProcessor).
This repo is intentionally tiny — one function, one log line — so you can focus on
the **trigger wiring** itself.

---

## What this function does

[`function_app.py`](function_app.py) declares a single function bound to the
`connector_trigger` decorator with a **strongly-typed**
`office365.ClientReceiveMessage` list from the Python SDK:

```python
@app.function_name(name="OnNewEmail")
@app.connector_trigger(arg_name="emails")
def on_new_email(emails: List[office365.ClientReceiveMessage]) -> None:
    logging.info("Received Microsoft 365 OnNewEmail trigger")
    for email in emails:
        logging.info("Email received from: %s", email.from_)
        logging.info("Email subject: %s", email.subject)
    # ...also dumps the full payload as JSON
```

When wired up to a Connector Namespace, the connectors platform handles the
webhook subscription, OAuth refresh, and retries — every new inbox message turns
into one invocation of this function.

---

## Prerequisites

- [Python 3.13](https://www.python.org/downloads/)
- [Azure Functions Core Tools v4](https://learn.microsoft.com/azure/azure-functions/functions-run-local#install-the-azure-functions-core-tools)
- [Azurite](https://learn.microsoft.com/azure/storage/common/storage-use-azurite)
  (local storage emulator) — required by the Functions host because
  [`local.settings.json`](local.settings.json) sets
  `AzureWebJobsStorage=UseDevelopmentStorage=true`. Install with
  `npm install -g azurite` or via the Azurite VS Code extension.
- [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli) with the
  preview `connector-namespace` extension installed (see step 0 below)
- [Microsoft Dev Tunnels CLI](https://learn.microsoft.com/azure/developer/dev-tunnels/get-started)
  (only needed for local testing against a real connection)
- An Azure subscription with access to **West Central US** (the current preview region)
- A Microsoft 365 / Office 365 mailbox you can authorize

### Recommended VS Code extensions

Open the folder in VS Code and accept the workspace recommendations
([`.vscode/extensions.json`](.vscode/extensions.json)):

| Extension | Purpose |
|---|---|
| Azure Functions | Run / debug / deploy the function |
| Python (`ms-python.python`) | Python IntelliSense and debugging |
| Pylance (`ms-python.vscode-pylance`) | Fast, feature-rich Python language support |
| REST Client (`humao.rest-client`) | Send the requests in [`test.http`](test.http) |
| Connector SDK IntelliSense ([VSIX](https://github.com/Azure/Connectors-NET-LSP/releases)) | Typed completions for SDK action methods (optional, early preview) |

---

## Create your local config files

Two files are gitignored (they can contain secrets). Copy the templates:

```bash
cp local.settings.json.sample local.settings.json
cp .env.sample .env
```

- **`local.settings.json`** — Functions host settings. Defaults work for local dev against Azurite. The devtunnel setup script appends any additional settings it needs (extension key, etc.) automatically.
- **`.env`** — used by [`test.http`](test.http) via the REST Client extension's `{{$dotenv VAR}}` syntax. Edit it after you run the setup script to set `HTTP_POST_CODE` (the function key).

---

## Step 0 — Install the `connector-namespace` CLI extension (once)

```bash
az extension add \
  --source https://github.com/anthonychu/azure-cli-extensions/releases/download/connector-namespace-0.1.0/connector_namespace-0.1.0-py2.py3-none-any.whl \
  --yes
az connector-namespace -h
```

---

## Set up the Python environment

```bash
python3.13 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

The `requirements.txt` pulls in:

```text
azure-functions>=2.2.0b3
azurefunctions-extensions-connectors
```

The `host.json` uses the **experimental extension bundle** (instead of NuGet) so
the connector trigger binding is automatically downloaded at host startup:

```json
{
  "extensionBundle": {
    "id": "Microsoft.Azure.Functions.ExtensionBundle.Experimental",
    "version": "[4.6.0, 5.0.0)"
  }
}
```

---

## Run locally (no real connection, just simulate webhook calls)

This is the fastest inner loop — start the host and POST a fake trigger payload
to the connector webhook endpoint.

1. **Start Azurite** (in a separate terminal, or via the VS Code Azurite extension's
   "Start" commands):

   ```bash
   azurite --silent --location .azurite
   ```

2. **Activate your virtual environment** (if not already active):

   ```bash
   source .venv/bin/activate
   ```

3. **Start the function host** in VS Code (`F5`) or from a terminal:

   ```bash
   func start
   ```

   You should see the function registered:

   ```text
   Functions:
       OnNewEmail: connectorTrigger
   ```

4. **Send a simulated trigger** using the REST Client extension and
   [`test.http`](test.http) — open the file and click **Send Request** above the
   first request. The function logs will show:

   ```text
   Received Microsoft 365 OnNewEmail trigger
   Email received from: someone@contoso.com
   Email subject: Hello from test.http
   ```

The webhook path used by the connector extension is:

```text
POST /runtime/webhooks/connector?functionName=OnNewEmail
```

For `func start` no `code` query parameter is required. For the deployed function
app you must append `&code=<connector_extension system key>` (see below).

---

## Run locally against a **real** Microsoft 365 connection (devtunnels)

The connectors service delivers events over the public internet, so to test your
local code against a real Office 365 inbox you need to expose port `7071` via a
public tunnel and point a Connector Namespace **trigger config** at that tunnel
URL.

The two scripts in [`scripts/`](scripts/) automate all of it:

| Script | Use it from |
|---|---|
| [`scripts/setup-devtunnel-connector.sh`](scripts/setup-devtunnel-connector.sh) | macOS / Linux / WSL (bash) |
| [`scripts/setup-devtunnel-connector.ps1`](scripts/setup-devtunnel-connector.ps1) | Windows (PowerShell) |

Both scripts:

1. Create (or reuse) an **anonymous Dev Tunnel** on port `7071` and start hosting it.
2. Create a **Connector Namespace** in `westcentralus` (if it doesn't exist).
3. Create an **Office 365 connection** and prompt you to complete OAuth consent
   in the browser.
4. Discover the `OnNewEmailV3` trigger operation on the connection.
5. Create a **trigger config** whose `callbackUrl` is your devtunnel URL +
   `/runtime/webhooks/connector?functionName=OnNewEmail`.

### Bash (macOS / Linux / WSL)

```bash
# In one terminal: activate venv and start the function host
source .venv/bin/activate
func start

# In another terminal: stand up the tunnel + connector wiring
./scripts/setup-devtunnel-connector.sh \
  --resource-group hello-m365-rg \
  --namespace      hello-m365-ns \
  --connection     office365-connection \
  --trigger-config on-new-email \
  --function-name  OnNewEmail
```

### PowerShell (Windows)

```powershell
# In one terminal: activate venv and start the function host
.venv\Scripts\activate
func start

# In another terminal: stand up the tunnel + connector wiring
./scripts/setup-devtunnel-connector.ps1 `
  -ResourceGroup  hello-m365-rg `
  -Namespace      hello-m365-ns `
  -Connection     office365-connection `
  -TriggerConfig  on-new-email `
  -FunctionName   OnNewEmail
```

> The first run will open a browser twice — once for `devtunnel user login`, once
> for the Office 365 OAuth consent.

Send yourself an email and watch the function log light up. To inspect what the
gateway delivered:

```bash
az connector-namespace trigger-config run list \
  -g hello-m365-rg --namespace-name hello-m365-ns \
  --trigger-config-name on-new-email -o table
```

When you're done, stop `func start` and stop the `devtunnel host` process
(`Ctrl+C` in both terminals).

---

## Deploy to Azure

This sample focuses on local development. To deploy:

1. Create a Flex Consumption Function App in any region supported by Functions
   (the **Connector Namespace** itself must be in `westcentralus`).
2. Use the VS Code Azure Functions extension or:

   ```bash
   func azure functionapp publish <your-function-app> --python
   ```

3. Fetch the connector extension system key and create a trigger config whose
   `callbackUrl` points at the deployed app:

   ```bash
   key=$(az functionapp keys list -g <rg> -n <app> \
           --query "systemKeys.connector_extension" -o tsv)

   az connector-namespace trigger-config create \
     --namespace-name hello-m365-ns \
     --resource-group hello-m365-rg \
     --name on-new-email \
     --available-connector office365 \
     --connection-name office365-connection \
     --operation-name OnNewEmailV3 \
     --parameter folderPath=Inbox \
     --callback-url "https://<app>.azurewebsites.net/runtime/webhooks/connector?functionName=OnNewEmail&code=$key"
   ```

See the [private preview instructions](https://gist.github.com/nzthiago/35799f1dc3b56d8f8915b1427e086c6c)
for the full walkthrough including Managed Identity access policies.

---

## Project layout

| File | Purpose |
|---|---|
| [function_app.py](function_app.py) | The single connector-triggered function (Python v2 model) |
| [requirements.txt](requirements.txt) | Python dependencies (`azure-functions`, `azurefunctions-extensions-connectors`) |
| [host.json](host.json) | Functions host config (experimental extension bundle) |
| [local.settings.json](local.settings.json) | Local app settings (not deployed) |
| [.python-version](.python-version) | Pins Python 3.13 for pyenv / VS Code |
| [test.http](test.http) | REST Client requests that simulate trigger callbacks |
| [scripts/setup-devtunnel-connector.sh](scripts/setup-devtunnel-connector.sh) | macOS/Linux devtunnel + `az connector-namespace` setup |
| [scripts/setup-devtunnel-connector.ps1](scripts/setup-devtunnel-connector.ps1) | Windows devtunnel + `az connector-namespace` setup |

---

## Feedback

Issues, ideas, missing connectors? Email
[`connectorsfuncprev@microsoft.com`](mailto:connectorsfuncprev@microsoft.com)
(see the [private preview gist](https://gist.github.com/nzthiago/35799f1dc3b56d8f8915b1427e086c6c)
for the full feedback channel and what the preview team is looking for).
