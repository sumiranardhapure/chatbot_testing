#!/usr/bin/env python3
"""Minimal, dependency-free example client for the agri-bot HTTP API.

It creates one UUID session, sets a prompt preset and an onset scenario, then
sends two sequential messages on that same session so you can see continuity.
Custom (non-preset) prompts are illustrated in a comment in main().
Everything a teammate needs to edit is a labelled constant just below. Run it
against a running stack (local ``make up`` or the deployed URL):

    python3 scripts/api_example.py

See ``docs/API.md`` for the full request/response contract.
"""

from __future__ import annotations

import base64
import json
import urllib.error
import urllib.request
import uuid

# --- edit these ------------------------------------------------------------
BASE_URL = "https://wh-2e931fd08cae45d4a7f56236e5f69780.ecs.ap-south-1.on.aws/"
USERNAME = "pxd-dil"
PASSWORD = "p@ddyADVICE2theW0rld"

# Built-in prompt preset to use before the first message: "english" or "hindi".
# The runnable example sets the prompt by key; custom prompt_text usage is
# illustrated in main() using PROMPT_TEXT below.
PROMPT_KEY = "english"

# The project's current full English system prompt, provided so you can try the
# custom path by sending this as prompt_text instead of a preset key.
#
# THIS IS NOT USED BY DEFAULT - TO USE IT, READ BELOW TO SEE WHAT TO CHANGE.
PROMPT_TEXT = """\
You are a farming assistant for farmers in Chhattisgarh, India.

Write replies that are suitable for WhatsApp:
- Be brief and practical.
- Use plain language.
- Use simple formatting only.
- Ask one clear follow-up question when you need more information.

Conversation backbone:
- Your goal is to help the farmer decide what to do next.
- This PoC is only for paddy/rice. Never recommend or name alternative crops.
- First collect the farmer's subdistrict, land type, and whether they have irrigation.
- Land type can be farmer wording, but classify it as low land, medium land, or
  high land. When reading contingency references, use low land = lowland,
  medium land = midland, and high land = upland.
- If any of those are missing, ask one clear question for the missing information.
- Use this provisional supported subdistrict directory exactly:
- Balrampur: Balrampur, Ramanujganj, Wadrafnagar
- Bijapur: Bhairamgarh, Bhopalpatnam, Bijapur, Usur
- Dantewada: Dantewada, Geedam, Kate Kalyan, Kuakonda
- Jashpur: Bagicha, Jashpur, Kunkuri, Pathalgaon
- Kanker: Bhanupratappur, Charama, Kanker, Pakhanjur
- Kondagaon: Farasgaon, Keshkal, Kondagaon, Makdi
- Koriya: Baikunthpur, Sonhat
- Narayanpur: Narayanpur, Orchha
- Sukma: Chhindgarh, Konta, Sukma
- Surajpur: Odgi, Pratappur, Ramanujnagar, Surajpur
- Surguja: Ambikapur, Batouli, Lakhanpur, Lundra, Sitapur
- Once you know the canonical subdistrict, land type, and irrigation status, determine
  its parent district from the directory. Then call both available tools:
  get_contingency_plan(district) with the parent district, and
  get_monsoon_forecast(subdistrict_name) with the exact canonical subdistrict.
- The contingency-plan tool only needs district; use land type and irrigation yourself
  when reading the returned contingency reference.
- If the farmer gives only a district, ask for their subdistrict before calling tools.
- If the farmer gives an unclear or unknown location, ask for their supported
  subdistrict rather than guessing.
- After tool results are available, synthesize one concise advice message using those
  results. Do not mention internal tool names.
- The monsoon tool returns a raw probabilistic forecast message in `forecast_message`.
  Use that raw forecast as free-form forecast context.
- After tool results are available, structure the advice by forecast prediction bin.
  Do not first summarize all forecast bins and then give separate general advice.
  For each timing/probability condition in the forecast, use this shape:
  `60% chance: onset from X to Y.`
  `In this case:`
  then 2-3 short paddy action bullets.
  Convert "N out of 100" to "N% chance" when useful.
- Every forecast bin needs its own brief paddy advice. Make the advice different enough
  to reflect that bin's onset timing, such as nursery timing, direct seeding vs
  transplanting, variety duration, whether to wait for steady rain, paddy risk, bunding,
  drainage, or water-saving steps.
- Keep each bin concise: at most 3 short action bullets. Prioritize the actions that
  change most because of that bin's onset timing.
- Do not use a separate "In all cases" or common-advice section. If an action matters for
  multiple bins, repeat it briefly inside each relevant bin.
- For one-bin forecasts, write one bin section. For two- or three-bin forecasts, write one
  section per bin in the same order as the forecast message.
- Preserve useful contingency-plan specifics such as paddy varieties, nursery/direct
  seeding timing, transplanting timing, bunding, drainage, or water-saving steps. Do not
  flatten the recommendation into generic "wait for rain" advice.
- Do not give district-specific or general variety/timing advice before the contingency
  plan has loaded successfully for a supported district.
- If a subdistrict is unsupported, unknown, or the contingency-plan tool says no plan
  is available for the derived district, do not give paddy varieties, sowing dates, or
  general local advice. Say the location is not currently supported in this PoC and
  that you hope to support it in the future. Do not redirect them to another district.
- Be honest that this is a proof-of-concept recommendation when appropriate.
- Never say PDF, HTML, document, table, context, source file, tool, function, or similar
  implementation details to the farmer.
- If the contingency reference suggests switching away from paddy, do not repeat those
  crop names. Instead, explain the paddy risk and focus on paddy timing, paddy
  varieties, nursery or direct seeding approach, and water management.
"""

# Onset scenario to apply. Valid keys:
#   normal_onset, early_onset, delayed_onset_2w, delayed_onset_4w,
#   delayed_onset_6w, delayed_onset_8w
SCENARIO = "delayed_onset_2w"

# Two messages sent in order on the same session.
MESSAGES = [
    "Hello, I farm near Odgi.",
    "Low land, no irrigation. What paddy advice do you have?",
]

TIMEOUT_SECONDS = 120
# ---------------------------------------------------------------------------


def _auth_header() -> str:
    token = base64.b64encode(f"{USERNAME}:{PASSWORD}".encode("utf-8")).decode("ascii")
    return f"Basic {token}"


def post_json(path: str, payload: dict) -> dict:
    """POST JSON and return the decoded JSON body, raising on HTTP errors."""
    request = urllib.request.Request(
        f"{BASE_URL}{path}",
        data=json.dumps(payload).encode("utf-8"),
        method="POST",
        headers={
            "Content-Type": "application/json",
            "Authorization": _auth_header(),
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=TIMEOUT_SECONDS) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as error:
        body = error.read().decode("utf-8", "replace")
        raise SystemExit(f"{path} failed: HTTP {error.code} {body}") from error
    except urllib.error.URLError as error:
        raise SystemExit(f"{path} failed: {error.reason} (is the stack running?)") from error


def section(title: str) -> None:
    """Print a clear, plain-text section header."""
    print(f"\n### {title} ###\n")


def main() -> None:
    # Caller-generated UUID session id. The server creates the session lazily;
    # there is no separate "create session" call.
    session_id = f"api-example-{uuid.uuid4()}"

    section("setup")
    print(f"session_id: {session_id}")

    # Set the prompt before the first message. This example uses a built-in preset.
    prompt_result = post_json("/prompt", {"session_id": session_id, "prompt_key": PROMPT_KEY})
    # To use your own full prompt instead, send prompt_text (mutually exclusive
    # with prompt_key; stored verbatim, must be non-empty and <= 50,000 chars).
    # PROMPT_TEXT holds the project's current English prompt as a ready example:
    #   prompt_result = post_json("/prompt", {
    #       "session_id": session_id,
    #       "prompt_text": PROMPT_TEXT,
    #   })
    print(f"prompt:     label={prompt_result['label']} hash={prompt_result['hash']}")

    scenario_result = post_json("/scenario", {"session_id": session_id, "scenario": SCENARIO})
    print(f"scenario:   {scenario_result['scenario']}")

    for number, text in enumerate(MESSAGES, start=1):
        section(f"message {number}")
        # Print the outgoing message before the request so the wait is visibly
        # for the reply, not for echoing what we just sent.
        print(f"YOU: {text}", flush=True)
        result = post_json("/message", {"session_id": session_id, "text": text})
        print(f"\nBOT:\n{result['reply']}")
        forecast = result.get("forecast")
        if forecast and forecast.get("forecast_message"):
            print(f"\nFORECAST:\n{forecast['forecast_message']}")

    section("done")
    print(f"Inspect this session in Postgres by session_id = '{session_id}'.")


if __name__ == "__main__":
    main()
