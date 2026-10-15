from common import *
from unittest.mock import patch
import pythoncom
from stage9_app.ownership import owned_office
calls=[]
def initialize():calls.append('initialize')
def uninitialize():calls.append('uninitialize')
RUN=OUT/('apartment_'+uuid4().hex[:8]);RUN.mkdir()
with patch.object(pythoncom,'CoInitialize',initialize),patch.object(pythoncom,'CoUninitialize',uninitialize):
    with owned_office(RUN/'office.json',supervised=True):
        for n in range(7):pythoncom.CoInitialize()
        for n in range(5):pythoncom.CoUninitialize()
    assert pythoncom.CoInitialize is initialize and pythoncom.CoUninitialize is uninitialize
assert calls==['initialize','uninitialize'],calls
state=read(next(RUN.glob('*COM_RELEASE_STATE*')))['data'];assert state['nested_apartment_initializations']==7 and state['deferred_apartment_releases']==5
save(RUN/'RESULT.json',{'status':'PASS','source_test_only':True,'calls':calls,'state':state})
print('STA ownership source regression PASS')
