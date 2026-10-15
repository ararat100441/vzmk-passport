from common import *
import sys
c.OUTPUT=Path(sys.argv[1])
c.recover_pending()
print("RECOVERY_COMPLETE")
