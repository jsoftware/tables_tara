NB. j504/j601 fix

cocurrent 'base'
load 'regex'

NB. =========================================================

3 : 0''
dat=. dat0=. fread jpath '~Addons/tables/tara/tara.ijs'
dat=. ('^\s*([xymnuv])\s*=\.\s*\1\..*\r?\n';'') rxrplc dat
dat=. ('coxtend ';'coinsert ') rxrplc dat
dat=. ('coxclass ';'coclass ') rxrplc dat
dat=. ('coxnew ';'conew ') rxrplc dat
dat fwrite jpath '~Addons/tables/tara/tara.ijs'
NB. tags j504
dat=. ('coxtend ';'(coinsert ] coextend) ') rxrplc dat0
dat=. ('coxclass ';'coclass ') rxrplc dat
dat=. ('coxnew ';'conew ') rxrplc dat
dat fwrite jpath '~Addons/tables/tara/tara-504.ijs'
i.0 0
)

