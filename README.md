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

Alternatively set `MICROSOFT_365_COPILOT_ACCESS_TOKEN`. The token is a short-lived Entra token intended for `substrate.office.com`, not a consumer Copilot token. Access tokens are typically valid for approximately 1 hour, after which re-authentication or token refresh may be required. Conversation IDs are managed automatically per Pi session and restored when you resume it. New sessions, forks, and tree navigation start separate conversations. The extension requests a new conversation, falling back to a generated ID if that endpoint returns HTTP 404.

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

## Local repository tools

Repository work uses Pi's enabled tools, not Copilot's remote sandbox. The provider
sends the host working directory and tool schemas to Copilot, converts explicit
`pi_tool_calls` JSON replies into Pi tool calls, and sends the resulting tool output
back on the next agent turn. Pi executes reads, edits, and commands locally through
its normal tool lifecycle.

File contents returned by tools are sent to Microsoft as conversation context.
The provider does not automatically upload the workspace or execute commands from
ordinary prose. With tools enabled, replies must contain either a `pi_tool_calls`
object or a `pi_final` object. The provider retries invalid replies up to twice,
then reports a protocol error instead of silently ending the turn. Unsupported
workspace-access claims are also retried unless a Pi tool reported a file-access
failure. With tools disabled, plain-text chat remains supported.

This is a prompt-based tool bridge, not a native Microsoft tool API; Copilot must
follow the tool protocol for local work to proceed.

## Prompt size and endpoint compatibility

There is no local character ceiling by default. The provider sends the full supplied
context, including system instructions, tool schemas, and tool results, without
size-based pruning. Any server size rejection is surfaced as a service error;
Microsoft's actual limit is not verified.

To opt into a local ceiling before starting Pi:

```sh
export MS365_COPILOT_MAX_TEXT_MESSAGE_LENGTH=12000
```

When configured, budgeting reserves headroom, drops older turns, and truncates
tool results as needed while keeping system instructions and tool definitions
intact. Unset `MS365_COPILOT_MAX_TEXT_MESSAGE_LENGTH` (including in `.env`) to
disable that local guardrail.

The protocol is undocumented and Microsoft can change it.

Development checks: `npm run typecheck`.
