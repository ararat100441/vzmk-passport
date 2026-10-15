from common import *
w=bootstrap('MAIN')
task('SOURCE','akz-settings-init',workspace=w)
d,p=task('SOURCE','akz-auto-check',workspace=w,vom=SETUP['vom'],variants='0,1,2,3')
generated,p=task('SOURCE','akz-auto-generate',workspace=w,preview=d['preview'],confirm_test=True)
save(OUT/'SOURCE_NATIVE_FINAL.json',{'status':'PASS','runtime':'SOURCE_NOT_EXE','workspace':str(w),'pairs':generated['pairs']})
