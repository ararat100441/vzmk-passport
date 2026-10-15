from common import *
checked,_,_=task('SOURCE_R5','akz-auto-check',vom=SETUP['vom'],variants='0,1,2,3')
save(OUT/'SOURCE_PREVIEW_R5.json',checked)
generated,result,sup=task('SOURCE_R5','akz-auto-generate',preview=checked['preview'],confirm_test=True)
save(OUT/'SOURCE_NATIVE_FINAL.json',{'status':'PASS','preview':checked['preview'],'result':str(result),'pairs':generated['pairs'],'runtime':'PYTHON_SOURCE_NOT_EXE'})
