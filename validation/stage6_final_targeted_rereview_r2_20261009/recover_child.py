from common_review import *
from akz_stage6 import core as c
c.OUTPUT=Path(sys.argv[1])
c.recover_pending()
print('FRESH_PROCESS_OWNED_PENDING_RECOVERED')
