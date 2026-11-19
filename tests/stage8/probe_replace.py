from pathlib import Path
import win32file
from contextlib import ExitStack
from stage8_closure.storage import frozen

r=Path('validation/stage8_20261010_01').resolve()
for i,(t,s,e) in enumerate([(0,0,0),(1,0,0),(0,1,0),(1,1,0),(1,1,1)]):
    a=r/str(i);b=r/(str(i)+'s');a.write_text('a');b.write_text('b')
    with ExitStack() as stack:
        if t:stack.enter_context(frozen([a],replaceable=a))
        if s:stack.enter_context(frozen([b],replaceable=b))
        x,y=str(a),str(b)
        if e:x,y='\\\\?\\'+x,'\\\\?\\'+y
        try:win32file.MoveFileEx(y,x,9);print(i,'PASS',len(y))
        except Exception as err:print(i,repr(err),len(y))
