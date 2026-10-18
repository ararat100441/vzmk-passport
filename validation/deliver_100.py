from pathlib import Path
import hashlib,json,shutil
root=Path(__file__).resolve().parents[1]
original=Path(r'C:\Users\arara\Desktop\ВЗМК рабочий стол\Работа\Документы\Pasport')
source=root/'dist/production100'
app=source/'VZMK-Passport-1.0.0-rc1'
shutil.copyfile(root/'docs/production_user_guide.txt',app/'ИНСТРУКЦИЯ.txt')
before=json.loads((root/'validation/production_100/BUILD_INPUTS.json').read_text(encoding='utf8'))
for path,h in before['files'].items():assert hashlib.sha256((root/path).read_bytes()).hexdigest()==h,path
destination=original/'stage9_dist_100_rc1';assert not destination.exists()
shutil.copytree(source,destination)
def hashes(folder):return {str(p.relative_to(folder)):hashlib.sha256(p.read_bytes()).hexdigest() for p in folder.rglob('*') if p.is_file()}
a,b=hashes(source),hashes(destination);assert a==b
exe=destination/'VZMK-Passport-1.0.0-rc1/VZMK-Passport-1.0.0-rc1.exe'
data={'version':'1.0.0-rc1','exe':str(exe),'sha256':hashlib.sha256(exe.read_bytes()).hexdigest(),'bytes':exe.stat().st_size,'files':len(a),'distribution_sha_match':True}
(root/'validation/production_100/DISTRIBUTION.json').write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf8')
print(json.dumps(data,ensure_ascii=False,indent=2))
