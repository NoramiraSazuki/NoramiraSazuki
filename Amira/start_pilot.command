#!/bin/zsh
# Start the Amira PDF pilot (local only). Double-click in Finder, or run ./start_pilot.command
# Keeps the Mac awake while running (caffeinate -i). Stop with Ctrl+C or the "Shut down" button.
cd "$(dirname "$0")" || exit 1
PY="../.venv/bin/python"
if [ ! -x "$PY" ]; then
  echo "Python environment not found at $PWD/$PY - see README section 'Local app'."
  read -k1 "?Press any key to close."
  exit 1
fi
caffeinate -i "$PY" -m pilot "$@"
status=$?
if [ $status -ne 0 ]; then read -k1 "?Pilot stopped with an error ($status). Press any key to close."; fi
exit $status
