NB. ---------------------------------------------------------
NB. Tests for writexlsheets
cocurrent 'base'
require '~addons/tables/tara/tara.ijs'

NB. Test data
x1=: i.3 5                  NB. int array
x2=: 9 15?.@$0              NB. float array
x3=: <"0 i.12 20            NB. boxed int array
x4=: 4 2$'abcd';'kdisd';'eiij asj' NB. boxed char array
x5=: 4 2$'abcd';54;'eiij';2;4.4 NB. boxed mixed array
x6=: 'No name, single char' NB. literal
x7=: 'data 1';'data 2'      NB. boxed char vector
x8=: 5 8                    NB. int vector
x9=: <"0 ]15.6 12.9 54.33   NB. boxed flt vector
x10=: 'Int array';x1      NB. With sheetnames
x11=: 'Flt array';x2
x12=: 'Box Int array';<x3
x13=: 'Box Chr array';<x4
x14=: 'Box Mix array';<x5
x15=: '';<x1              NB. empty sheetname
x16=: ,.x2;<x3   NB. data column- flt and boxed int arrays
x17=: (<x9),(<x7),(<x8),(<x5),(<x1),(<x4),x16 NB. data column
x18=: x10 ,: 'Literal';x6
x19=: (x10,x11,x12,x13,x14,:x15)
x20=: (,.5#<''),(4#<''),"1 x3 NB. offset topleft corner by 5 rows and 4 cols
char=:('per Angusta';'ad Augusta'),~;:'Lorem ipsum dolor sit consectetur'
idx =: 2 2;2 3;2 4;2 5;3 2;3 3;3 5;4 2;4 4;4 5;5 2;5 3;5 4;5 5
x21=: 'Diff Mix array';<(14$char) idx} <"0 i.8 9


fnme=: jpath '~temp/tarawsht'
suff=: '.xls'

NB.Tests
x1 writexlsheets fnme,'1',suff
x2 writexlsheets fnme,'2',suff
x3 writexlsheets fnme,'3',suff
x4 writexlsheets fnme,'4',suff
x5 writexlsheets fnme,'5',suff
x6 writexlsheets fnme,'6',suff
x7 writexlsheets fnme,'7',suff
x8 writexlsheets fnme,'8',suff
x9 writexlsheets fnme,'9',suff
x10 writexlsheets fnme,'10',suff
x11 writexlsheets fnme,'11',suff
x12 writexlsheets fnme,'12',suff
x13 writexlsheets fnme,'13',suff
x14 writexlsheets fnme,'14',suff
x15 writexlsheets fnme,'15',suff
x16 writexlsheets fnme,'16',suff
x17 writexlsheets fnme,'17',suff
x18 writexlsheets fnme,'18',suff
x19 writexlsheets fnme,'19',suff
x20 writexlsheets fnme,'20',suff
x21 writexlsheets fnme,'21',suff