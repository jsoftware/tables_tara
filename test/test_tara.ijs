NB. Tests for tara

Note 'To run all tests:'
  load 'tables/tara'
  load 'tables/tara/test/test_tara'
)

loc=. 3 : '> (4!:4 <''y'') { 4!:3 $0'
PATH=. getpath_j_ loc''

NB. -------------------------------------------------------
NB. scripts for testing

load PATH,'taratest.ijs'
load PATH,'taratest2.ijs'
