from pathlib import Path
root=Path(SPECPATH)
packages=['stage9_app','passport_app','profiles_stage4','uzk_stage5','akz_stage6','stage7_quality','stage8_closure']
datas=[]
for package in packages:
    datas += [(str(p),package) for p in (root/package).glob('*.py')]
datas += [(str(root/'stage8.py'),'.'),(str(root/'stage9.py'),'.'),(str(root/'stage7_quality/render_emf.ps1'),'stage7_quality')]
a=Analysis([str(root/'stage9.py')],pathex=[str(root)],binaries=[],datas=datas,
    hiddenimports=['win32timezone','pythoncom','pywintypes','win32com.client','win32com.client.dynamic','win32com.client.gencache','win32api','win32file','win32con','win32security','win32process','win32event','olefile'],
    hookspath=[],hooksconfig={},runtime_hooks=[],excludes=['pytest','matplotlib','pandas','numpy'],noarchive=False)
pyz=PYZ(a.pure)
exe=EXE(pyz,a.scripts,[],exclude_binaries=True,name='VZMK-Passport-0.9.16-DEVTEST',debug=False,bootloader_ignore_signals=False,strip=False,upx=False,console=False)
coll=COLLECT(exe,a.binaries,a.datas,strip=False,upx=False,name='VZMK-Passport-0.9.16-DEVTEST')
