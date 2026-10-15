from pathlib import Path
import sys
from akz_stage6 import core as m
m.OUTPUT=Path(sys.argv[1])
m.recover_pending()
