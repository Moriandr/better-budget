#!/usr/bin/env python3
"""Builds dist/Better Budget.gs (the single file users paste into Apps Script)
from the modules in src/modules, and copies the HTML page and manifest next to it.

Usage:  python3 tools/build.py
"""
import os, shutil

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MODULES = ["Custom Menu", "Settings and Setup", "Styling", "First Run", "Record sidebar",
           "Monthly Reports", "New Month", "Dashboard", "Weekly Report", "About"]

HEADER = """// ============================================================================
// BETTER BUDGET - COMPLETE SCRIPT
// ============================================================================
// This single file contains the whole tool. Sections, in order:
//   1. Custom Menu          2. Settings and Setup    3. Styling
//   4. First Run            5. Record sidebar        6. Monthly Reports
//   7. New Month            8. Dashboard             9. Weekly Report
//  10. About
// The entry form lives in the separate SidebarHTML file.
// ============================================================================
"""

def main():
    parts = [HEADER]
    for i, name in enumerate(MODULES, 1):
        with open(os.path.join(ROOT, "src", "modules", name + ".gs"), encoding="utf-8") as f:
            text = f.read().strip("\n")
        bar = "=" * 76
        parts.append(f"\n\n// {bar}\n// SECTION {i} OF {len(MODULES)}: {name.upper()}\n// {bar}\n\n{text}\n")
    dist = os.path.join(ROOT, "dist")
    os.makedirs(dist, exist_ok=True)
    with open(os.path.join(dist, "Better Budget.gs"), "w", encoding="utf-8") as f:
        f.write("".join(parts))
    for extra in ("SidebarHTML.html", "appsscript.json"):
        shutil.copy(os.path.join(ROOT, "src", extra), dist)
    shutil.copy(os.path.join(ROOT, "docs", "INSTALL.txt"), os.path.join(dist, "README.txt"))
    print("Built dist/ (Better Budget.gs, SidebarHTML.html, appsscript.json, README.txt)")

if __name__ == "__main__":
    main()
