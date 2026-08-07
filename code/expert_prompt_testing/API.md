# HTTP API

This is the practical guide for driving the agri-bot over HTTP — from the
browser testers, the bulk runner, `scripts/api_example.py`, or your own client.
It is a small JSON API behind HTTP Basic Auth. There is no SDK and no OpenAPI
spec; everything you need is below.

For the internal design behind these routes, see `docs/ARCHITECTURE.md`.

## Base URL and auth

- **Base URL:** `http://localhost:8000` locally, or the deployed HTTPS URL.
- **Auth:** HTTP Basic Auth on every route except `/health`, `/version`, and the
  favicons. Use the shared `BASIC_AUTH_USER` / `BASIC_AUTH_PASS` credentials.
- **Content type:** send `Content-Type: application/json`; bodies are JSON.

```bash
curl -s -u "$BASIC_AUTH_USER:$BASIC_AUTH_PASS" \
  -H "Content-Type: application/json" \
  -d '{"session_id":"demo-1","text":"hello"}' \
  http://localhost:8000/message
```

A missing or wrong credential returns `401` with a `WWW-Authenticate` header.

## Sessions are caller-generated UUIDs

There is **no create-session endpoint**. You invent a session id (use a UUID to
avoid collisions) and pass it on every request. The server creates the row
lazily on first use. The same id continues a conversation; a brand-new id starts
a fresh one. Prior sessions are never deleted or reused — "starting a new
session" just means generating a new id.

```
session_id = "client-" + uuid4()      # e.g. client-9f1c...e2
```

## Setup before the first message

Prompt and scenario are **locked once the first `/message` is sent** (the lock
triggers as soon as the session has any conversation history). Set them first.

### `POST /prompt`

Send **exactly one** of `prompt_key` or `prompt_text` (never both, never
neither).

- `prompt_key` — a preset lookup tag: `"english"` or `"hindi"`. It is resolved
  to full prompt text and a hash at request time and is **never stored or
  logged**.
- `prompt_text` — a full custom system prompt, stored verbatim. Must be
  non-empty and at most 50,000 characters.

Preset request/response:

```bash
curl -s -u "$BASIC_AUTH_USER:$BASIC_AUTH_PASS" -H "Content-Type: application/json" \
  -d '{"session_id":"demo-1","prompt_key":"hindi"}' \
  http://localhost:8000/prompt
# {"status":"ok","prompt_key":"hindi","label":"Hindi","hash":"<12-hex>"}
```

Custom prompt request/response (the response echoes only a label and the
computed hash — never the prompt text):

```bash
curl -s -u "$BASIC_AUTH_USER:$BASIC_AUTH_PASS" -H "Content-Type: application/json" \
  -d '{"session_id":"demo-1","prompt_text":"You are a terse paddy bot..."}' \
  http://localhost:8000/prompt
# {"status":"ok","label":"custom","hash":"<12-hex>"}
```

The `hash` is the first 12 hex chars of the SHA-256 of the exact prompt text.
Identical prompt text always yields the same hash, so you can group sessions by
which prompt they ran (see "Inspecting stored sessions").

### `POST /scenario`

Selects the mocked monsoon-onset scenario for the session.

```bash
curl -s -u "$BASIC_AUTH_USER:$BASIC_AUTH_PASS" -H "Content-Type: application/json" \
  -d '{"session_id":"demo-1","scenario":"delayed_onset_4w"}' \
  http://localhost:8000/scenario
# {"status":"ok","scenario":"delayed_onset_4w"}
```

Valid scenarios: `normal_onset`, `early_onset`, `delayed_onset_2w`,
`delayed_onset_4w`, `delayed_onset_6w`, `delayed_onset_8w`.

## Sending messages

### `POST /message`

```bash
curl -s -u "$BASIC_AUTH_USER:$BASIC_AUTH_PASS" -H "Content-Type: application/json" \
  -d '{"session_id":"demo-1","text":"Odgi, low land, no irrigation"}' \
  http://localhost:8000/message
```

Response:

```json
{
  "reply": "…the bot's WhatsApp-style reply…",
  "forecast": {
    "forecast_message": "…raw forecast text, or null until the tool runs…"
  }
}
```

**Continuation:** send another `/message` with the **same** `session_id` to
continue the conversation; the server carries prior context server-side. The
first `/message` locks the prompt and scenario for that session.

**Starting a new session:** generate a new `session_id` and (optionally) set a
new prompt/scenario before its first message.

## Errors

| Status | When |
|--------|------|
| `400` | Missing/blank `session_id`; missing `text`; `/scenario` with an unknown scenario; `/prompt` with both or neither of `prompt_key`/`prompt_text`, an invalid `prompt_key`, or empty/over-long `prompt_text`. |
| `401` | Missing or wrong Basic Auth credentials. |
| `409` | `/prompt` or `/scenario` after the session's first message (setup is locked). |

Error bodies are `{"error": "<message>"}`.

## Inspecting stored sessions

Durable state lives in Postgres. The `sessions` table holds one row per
`session_id` with the exact `prompt`, its `prompt_hash`, the `scenario`,
`model_config`, and the `history`. The `session_logs` table holds one row per
`session_id` with a `turns` array (response ids, model calls, tool
invocations). The prompt text and hash are **not** copied into `session_logs`.

Filter sessions by which prompt they ran, and join to their interaction trace:

```sql
-- All sessions that ran a specific prompt (by hash), with their turn count.
SELECT s.session_id,
       s.prompt_hash,
       s.scenario,
       jsonb_array_length(l.turns) AS turns
FROM sessions s
LEFT JOIN session_logs l USING (session_id)
WHERE s.prompt_hash = '<12-hex-hash>'
ORDER BY s.created_at;
```

To find the hash for a given prompt without running anything:

```python
import hashlib
print(hashlib.sha256(prompt_text.encode("utf-8")).hexdigest()[:12])
```

## Worked example

`scripts/api_example.py` is a dependency-free client that creates a UUID
session, sets a custom prompt and scenario, sends two sequential messages on the
same session, and prints the whole exchange. Edit the labelled constants at the
top (URL, credentials, prompt, scenario, messages, timeout) and run:

```bash
python3 scripts/api_example.py
```
