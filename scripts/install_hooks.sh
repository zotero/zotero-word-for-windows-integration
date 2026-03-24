#!/bin/bash

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
HOOKS_DIR="$SCRIPT_DIR/hooks"

git config core.hooksPath "$HOOKS_DIR"
echo "Git hooks path set to $HOOKS_DIR"
