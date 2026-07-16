#!/bin/bash
# Wrapper for launchd — captures errors before run_hourly_publish.sh can log them
LOG="/Users/apenkrat/ACC-Portfolio-Review/export.log"
echo "=== launchd wrapper start $(date) ===" >> "$LOG" 2>&1
export PATH="/opt/homebrew/bin:/opt/homebrew/sbin:/Library/Frameworks/Python.framework/Versions/3.14/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin"
export HOME="/Users/apenkrat"
cd /Users/apenkrat/ACC-Portfolio-Review 2>>"$LOG" || { echo "cd failed" >> "$LOG"; exit 1; }
/Users/apenkrat/ACC-Portfolio-Review/run_hourly_publish.sh >> "$LOG" 2>&1
echo "=== launchd wrapper exit $? ===" >> "$LOG"
