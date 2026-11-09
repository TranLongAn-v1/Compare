import os
import sys

base_dir = getattr(sys, "_MEIPASS", os.path.abspath("."))
tcl_candidates = [
    os.path.join(base_dir, "_tcl_data"),
    os.path.join(base_dir, "tcl8.6"),
]
tk_candidates = [
    os.path.join(base_dir, "_tk_data"),
    os.path.join(base_dir, "tk8.6"),
]

for candidate in tcl_candidates:
    if os.path.isdir(candidate):
        os.environ["TCL_LIBRARY"] = candidate
        break

for candidate in tk_candidates:
    if os.path.isdir(candidate):
        os.environ["TK_LIBRARY"] = candidate
        break
