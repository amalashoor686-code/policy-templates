import anthropic

# Credentials come from workload identity federation. The client reads
# ANTHROPIC_FEDERATION_RULE_ID, ANTHROPIC_ORGANIZATION_ID,
# ANTHROPIC_SERVICE_ACCOUNT_ID, ANTHROPIC_WORKSPACE_ID (optional for
# single-workspace rules) and ANTHROPIC_IDENTITY_TOKEN_FILE from the
# environment. No API key is used.
client = anthropic.Anthropic()

message = client.messages.create(
    model="claude-opus-5-5",
    max_tokens=1024,
    messages=[{"role": "user", "content": "Hello, Claude"}],
)
print(next(block.text for block in message.content if block.type == "text"))
