#!/bin/bash
# Restores real secret values from placeholders on checkout.
# Reads .secrets file for key=value pairs.

SECRETS_FILE="$(git rev-parse --show-toplevel)/.secrets"
INPUT=$(cat)

if [ -f "$SECRETS_FILE" ]; then
  while IFS='=' read -r key value; do
    [ -z "$key" ] && continue
    [[ "$key" =~ ^# ]] && continue
    INPUT=$(echo "$INPUT" | sed "s|YOUR_${key}|$value|g")
  done < "$SECRETS_FILE"
fi

echo "$INPUT"
