from common import *
settings,_,_=task('SOURCE','akz-settings-init')
checked,preview_result,_=task('SOURCE','akz-auto-check',vom=SETUP['vom'],variants='0,1,2,3')
generated,result,sup=task('SOURCE','akz-auto-generate',preview=checked['preview'],confirm_test=True)
save(OUT/'SOURCE_NATIVE_FINAL.json',{'status':'PASS','settings':settings,'preview':checked['preview'],'result':str(result),'pairs':generated['pairs'],'runtime':'PYTHON_SOURCE_NOT_EXE'})
