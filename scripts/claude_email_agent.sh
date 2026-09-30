#!/usr/bin/env bash
# Runs Claude with a send_email tool over plain curl and executes the tool
# calls itself. Usage: claude_email_agent.sh "<prompt>"
#
# Email is sent from an Outlook / Microsoft 365 mailbox via Microsoft Graph.
#
# Required env: ACCESS_TOKEN (Anthropic bearer token),
#   AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET (Entra app with the
#   Mail.Send application permission), EMAIL_FROM (the sending mailbox),
#   EMAIL_ALLOWED_RECIPIENTS (comma-separated)
# Optional env: MODEL (default claude-sonnet-5-5), MAX_TURNS (default 5)
set -euo pipefail

prompt=${1:?usage: $0 "<prompt>"}
model=${MODEL:-claude-sonnet-5-5}
max_turns=${MAX_TURNS:-5}

tools='[{
  "name": "send_email",
  "description": "Send a plain-text email. Recipients must be on the operator-approved allowlist; other addresses are rejected.",
  "strict": true,
  "input_schema": {
    "type": "object",
    "properties": {
      "to": {"type": "string", "description": "Recipient email address"},
      "subject": {"type": "string"},
      "body": {"type": "string", "description": "Plain-text message body"}
    },
    "required": ["to", "subject", "body"],
    "additionalProperties": false
  }
}]'

# Prints nothing and returns 0 on success; prints the reason and returns 1 on failure.
send_email() {
  "$(dirname "$0")/outlook_send.sh" \
    "$(jq -r .to <<<"$1")" "$(jq -r .subject <<<"$1")" "$(jq -r .body <<<"$1")"
}

messages=$(jq -n --arg p "$prompt" '[{role: "user", content: $p}]')

for ((turn = 1; turn <= max_turns; turn++)); do
  payload=$(jq -n --arg m "$model" --argjson msgs "$messages" --argjson tools "$tools" \
    '{model: $m, max_tokens: 16000, thinking: {type: "adaptive"},
      output_config: {effort: "high"}, tools: $tools, messages: $msgs}')
  resp=$(curl -sS --max-time 600 https://api.anthropic.com/v1/messages \
    -H "authorization: Bearer $ACCESS_TOKEN" \
    -H "anthropic-version: 2023-06-01" \
    -H "content-type: application/json" -d "$payload")

  if ! jq -e '.content' >/dev/null <<<"$resp"; then
    echo "API error: $(jq -c '.error // .' <<<"$resp")" >&2
    exit 1
  fi

  # Append the assistant turn unchanged (thinking blocks must be echoed back).
  messages=$(jq --argjson r "$resp" '. + [{role: "assistant", content: $r.content}]' <<<"$messages")

  stop=$(jq -r .stop_reason <<<"$resp")
  if [[ $stop != tool_use ]]; then
    [[ $stop == refusal || $stop == max_tokens ]] && echo "::warning::stop_reason=$stop"
    jq -r '.content[] | select(.type == "text") | .text' <<<"$resp"
    exit 0
  fi

  results='[]'
  while IFS= read -r block; do
    id=$(jq -r .id <<<"$block")
    name=$(jq -r .name <<<"$block")
    input=$(jq -c .input <<<"$block")
    is_error=false
    if [[ $name == send_email ]]; then
      if out=$(send_email "$input"); then
        result="Email sent to $(jq -r .to <<<"$input")."
      else
        result="Failed to send email: $out"; is_error=true
      fi
    else
      result="Unknown tool: $name"; is_error=true
    fi
    echo "tool_use $name: $result" >&2
    results=$(jq --arg id "$id" --arg c "$result" --argjson e "$is_error" \
      '. + [{type: "tool_result", tool_use_id: $id, content: $c, is_error: $e}]' <<<"$results")
  done < <(jq -c '.content[] | select(.type == "tool_use")' <<<"$resp")

  messages=$(jq --argjson r "$results" '. + [{role: "user", content: $r}]' <<<"$messages")
done

echo "Stopped after $max_turns turns without a final answer." >&2
exit 1
