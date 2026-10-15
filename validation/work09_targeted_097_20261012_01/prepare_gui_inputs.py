from common import *
import shutil
# A new isolated folder makes native file selection possible without text input.
folder=Path.home()/'Documents/!V9W09_7c3bb1e3';folder.mkdir()
p=folder/'TEST_VOM.xls';shutil.copyfile(VOM,p);assert sha(p)==sha(VOM)
save(OUT/'GUI_NAVIGATION_FIXTURE.json',{'root':str(folder),'vom':str(p),'sha256':sha(p),'source':str(VOM),'purpose':'isolated GUI input selected through the visible native file list; no user files overwritten'})
print(folder,flush=True)
