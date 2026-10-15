from common import *
checked,_,_=task('SOURCE_FORMAT','akz-auto-check',vom=SETUP['vom'],variants='2')
generated,result,sup=task('SOURCE_FORMAT','akz-auto-generate',preview=checked['preview'],confirm_test=True)
save(OUT/'SOURCE_FORMAT_FINAL.json',{'status':'PASS','result':str(result),'pairs':generated['pairs'],'runtime':'PYTHON_SOURCE_NOT_EXE'})
