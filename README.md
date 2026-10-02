# pi-ms365-copilot

Microsoft 365 Copilot provider extension for Pi.

## Install

Install directly from GitHub:

```bash
pi install https://github.com/ilpianista/pi-ms365-copilot
```

You can also pin a ref:

```bash
pi install https://github.com/ilpianista/pi-ms365-copilot@master
```

After install, restart `pi` and select the `ms365-copilot/copilot` model.

For interactive use, log in once from inside `pi`:

```text
/login
```

Then select **Microsoft 365 Copilot** from the OAuth provider list. (Direct `/login ms365-copilot` is currently treated as a regular prompt sent to the agent by pi. Use bare `/login` instead.)

Alternatively set `MICROSOFT_365_COPILOT_ACCESS_TOKEN`. The token is a short-lived Entra token intended for `https://substrate.office.com/sydney`, not a consumer Copilot token. Conversation IDs are managed automatically per Pi session and restored when you resume it. New sessions, forks, and tree navigation start separate conversations. The extension requests a new conversation, falling back to a generated ID if that endpoint returns HTTP 404.

## Models / tones

Select a tone using Pi's `/model` picker. All IDs below use the `ms365-copilot/` provider prefix:

| Model ID | Choice | Copilot wire `tone` |
| --- | --- | --- |
| `copilot` | Auto (existing/default ID) | `Magic` |
| `copilot-quick` | Quick | `Chat` |
| `copilot-deeper` | Think Deeper | `Reasoning` |
| `copilot-sol-quick` | SOL Quick | `Gpt_5_6_Chat` |
| `copilot-sol-think` | SOL Think | `Gpt_5_6_Reasoning` |

These aren't guarantees of a particular underlying model version. Availability and actual routing depend on Microsoft and your account.

## Extracting the access token

After starting a chat at <https://copilot.cloud.microsoft/>, open DevTools, find the WebSocket request to `wss://substrate.office.com/m365Copilot/Chathub/`, and copy its complete URL.

Extract the `access_token` query parameter from the URL or execute the following command:

```sh
wl-paste | python3 -c 'import sys; from urllib.parse import parse_qs, urlsplit; print(parse_qs(urlsplit(sys.stdin.read().strip()).query)["access_token"][0])'
```

## Prompt size and endpoint compatibility

The extension uses a local 8000-character message ceiling by default. This is a local guardrail, not a verified Microsoft server limit. Prompt budgeting still reserves headroom and keeps system instructions and tool definitions intact; large Pi contexts may need fewer tools or shorter instructions.

To configure the local ceiling before starting Pi:

```sh
export MS365_COPILOT_MAX_TEXT_MESSAGE_LENGTH=12000
```

The protocol is undocumented and Microsoft can change it.
