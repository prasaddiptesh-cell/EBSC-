#!/usr/bin/env python3
"""Entry point used for both Python development and the packaged Windows executable."""
from __future__ import annotations
import sys

def main() -> int:
    if len(sys.argv) >= 2 and sys.argv[1] == "--bridge-command":
        from bridge import main as bridge_main
        sys.argv = ["bridge.py", *sys.argv[2:]]
        return bridge_main()
    from bridge_ui import BridgeUI
    BridgeUI().mainloop()
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
