#!/bin/env bash

printf '\n## TREE ##\n'
find src -type f | sort
printf '\n## FILES ##\n'
find src -type f -name '*.py' | sort | while read -r f; do
  printf '\n### FILE: %s ###\n' "$f"
  cat "$f"
  printf '\n'
done
