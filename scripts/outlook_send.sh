#!/usr/bin/env bash
# Sends a plain-text email from an Outlook / Microsoft 365 mailbox via
# Microsoft Graph. Usage: outlook_send.sh <to> <subject> <body>
#
# Required env: AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET (Entra app
#   with the Mail.Send application permission), EMAIL_FROM (sending mailbox),
#   EMAIL_ALLOWED_RECIPIENTS (comma-separated; anything else is refused)
# Prints the failure reason to stdout and exits 1 on failure.
set -euo pipefail

to=$(tr -d '\r\n' <<<"${1:?usage: $0 <to> <subject> <body>}")
subject=$(tr -d '\r\n' <<<"${2:?missing subject}")
body=${3:?missing body}

allowed=false
IFS=',' read -ra list <<<"$EMAIL_ALLOWED_RECIPIENTS"
for addr in "${list[@]}"; do
  addr=$(tr -d '[:space:]' <<<"$addr")
  if [[ -n $addr && ${addr,,} == "${to,,}" ]]; then allowed=true; fi
done
if [[ $allowed != true ]]; then
  echo "Recipient $to is not on the allowlist."
  exit 1
fi

graph_token=$(curl -sS -X POST \
  "https://login.microsoftonline.com/$AZURE_TENANT_ID/oauth2/v2.0/token" \
  --data-urlencode "client_id=$AZURE_CLIENT_ID" \
  --data-urlencode "client_secret=$AZURE_CLIENT_SECRET" \
  --data-urlencode "scope=https://graph.microsoft.com/.default" \
  --data-urlencode "grant_type=client_credentials" | jq -er .access_token) || {
  echo "Could not get a Microsoft Graph token."
  exit 1
}

payload=$(jq -n --arg to "$to" --arg s "$subject" --arg b "$body" \
  '{message: {subject: $s, body: {contentType: "Text", content: $b},
              toRecipients: [{emailAddress: {address: $to}}]},
    saveToSentItems: true}')
out=$(mktemp)
status=$(curl -sS -o "$out" -w '%{http_code}' -X POST \
  "https://graph.microsoft.com/v1.0/users/$EMAIL_FROM/sendMail" \
  -H "authorization: Bearer $graph_token" \
  -H "content-type: application/json" -d "$payload") || exit 1
if [[ $status != 202 ]]; then
  echo "Graph sendMail returned HTTP $status: $(jq -r '.error.message // empty' "$out" 2>/dev/null)"
  exit 1
fi
