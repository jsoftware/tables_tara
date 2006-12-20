NB. j504/j601 fix

cocurrent 'base'
load 'regex'

NB. =========================================================

3 : 0''
if. 504 < 0&". 'j'-.~4{.9!:14 '' do.
  dat=. fread jpath '~Addons/tables/tara/tara.ijs'
  dat=. ('^\s*([xymnuv])\s*=\.\s*\1\..*\r?\n|^\s*coxinsert .*\r?\n';'') rxrplc dat
  dat=. ('coxclass ';'coclass ') rxrplc dat
  dat=. ('coxnew ';'conew ') rxrplc dat
  dat fwrite jpath '~Addons/tables/tara/tara.ijs'
else.
  dat=. fread jpath '~Addons/tables/tara/tara.ijs'
  dat=. ('coxinsert ';'coinsert ') rxrplc dat
  dat=. ('coxclass ';'coclass ') rxrplc dat
  dat=. ('coxnew ';'conew ') rxrplc dat
  dat fwrite jpath '~Addons/tables/tara/tara.ijs'
end.
i.0 0
)

