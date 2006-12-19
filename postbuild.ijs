NB. j504/j601 fix

cocurrent 'base'

NB. =========================================================

colib=: 0 : 0
18!:4 <'z'
coxclass=: 18!:4 @ boxxopen

NB.*conew v create object
coxnew=: 3 : 0
c=. <y.
obj=. cocreate''
coinsert__obj c
COCREATOR__obj=: coname''
obj
:
w=. coxnew y.
create__w x.
w
()

)

3 : 0''
if. 504 < 0&". 'j'-.~4{.9!:14 '' do.
  dat=. freads jpath '~Addons/tables/tara/tara.ijs'
  dat=. dat stringreplace~ 'coxclass ''' ; 'coclass ''' ; 'coxnew '''; 'conew ''' ; (LF, 'y=. y.', LF) ; LF ; (LF, 'y=. y. [ x=. x.', LF) ; LF
  dat fwrites jpath '~Addons/tables/tara/tara.ijs'
else.
  dat=. freads jpath '~Addons/tables/tara/tara.ijs'
  dat=. dat stringreplace~ 'coxclass ''' ; 'coclass ''' ; 'coxnew '''; 'conew '''
  dat fwrites jpath '~Addons/tables/tara/tara.ijs'
end.
i.0 0
)
