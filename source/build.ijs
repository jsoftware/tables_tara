NB. build.ijs

writesourcex_jp_ '~Addons/tables/tara/source';'~Addons/tables/tara/tara.ijs'

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

f=. 3 : 0
mkdir_j_ jpath '~addons/tables/tara/test/'
(jpath '~addons/tables/tara/',y) fcopynew ::] jpath '~Addons/tables/tara/',y
)

f 'tara.ijs'
f 'tara.ijt'
f 'history.txt'
f 'taradoc.txt'
f 'test/dora.png'
f 'test/taratest.ijs'
f 'test/taratest2.ijs'
f 'test/test.xls'
