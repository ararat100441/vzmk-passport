from pathlib import Path
from production_app import VERSION
root=Path(SPECPATH)
packages=['production_app','stage9_app','passport_app','profiles_stage4','uzk_stage5','akz_stage6','stage7_quality','stage8_closure']
datas=[(str(p),package) for package in packages for p in (root/package).glob('*.py')]
datas += [(str(root/'docs/production_user_guide.txt'),'.')]
a=Analysis([str(root/'production.py')],pathex=[str(root)],binaries=[],datas=datas,
    hiddenimports=['pypdf','win32timezone','pythoncom','pywintypes','win32com.client','win32com.client.dynamic','win32com.client.gencache','win32api','win32file','win32con','win32security','win32process','win32event','olefile'],
    hookspath=[],hooksconfig={},runtime_hooks=[],excludes=['pytest','matplotlib','pandas','numpy'],noarchive=False)
pyz=PYZ(a.pure)
exe=EXE(pyz,a.scripts,[],exclude_binaries=True,name='VZMK-Passport-'+VERSION,debug=False,bootloader_ignore_signals=False,strip=False,upx=False,console=False)
coll=COLLECT(exe,a.binaries,a.datas,strip=False,upx=False,name='VZMK-Passport-'+VERSION)
