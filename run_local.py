"""Local launcher; also supports dependencies installed into .local-deps."""
import runpy
import sys
from pathlib import Path

root = Path(__file__).resolve().parent
local_deps = root / ".local-deps"
if local_deps.is_dir():
    sys.path.insert(0, str(local_deps))
sys.path.insert(0, str(root))
sys.argv = ["streamlit", "run", str(root / "app.py"), "--global.developmentMode", "false", *sys.argv[1:]]
runpy.run_module("streamlit", run_name="__main__")
