coclass 'oleutlfcn'
oledate2local=: 3 : 0
86400000* _72682+86400%~10000000%~(8#256)#. a.i.y
)
localdate2ole=: 3 : 0
a.{~(8#256)#: 10000000*x:86400*(y%86400000)+72682
)
Note=: 3 : '0 0 $ 0 : 0' : [
bitand=: 17 b.
bitxor=: 22 b.
bitor=: 23 b.
bitneg=: 26 b.
bitrot=: 32 b.
bitshl=: 33 b.
bitsha=: 34 b.
bigendian=: ({.a.)={. 1&(3!:4) 1  
toBYTE=: {&a.@(256&|)
fromBYTE=: a.&i.
toWORDm=: 1&(3!:4)@:<.
toDWORDm=: 2&(3!:4)@:<.
toucodem=: ''&,@(1&(3!:4))@(3&u:)@u:
toDoublem=: 2&(3!:5)
fromWORDm=: _1&(3!:4)
fromDWORDm=: _2&(3!:4)
fromucodem=: 6&u:
fromDoublem=: _2&(3!:5)
toWORDr=: ,@:(|."1)@(_2: ]\ 1&(3!:4)@:<.)
toDWORDr=: ,@:(|."1)@(_4: ]\ 2&(3!:4)@:<.)
toucoder=: ''&,@:,@:(|."1@(_2: ]\ 1&(3!:4)))@(3&u:)@u:
toDoubler=: ,@:(|."1)@(_8: ]\ 2&(3!:5))
fromWORDr=: _1&(3!:4)@:,@:(|."1)@(_2&(]\))
fromDWORDr=: _2&(3!:4)@:,@:(|."1)@(_4&(]\))
fromucoder=: 6&u:@:,@:(|."1)@(_2&(]\))
fromDoubler=: _2&(3!:5)@:,@:(|."1)@(_8&(]\))
toWORD0=: toWORDm`toWORDr@.bigendian f.
toDWORD0=: toDWORDm`toDWORDr@.bigendian f.
toucode0=: toucodem`toucoder@.bigendian f.
toDouble0=: toDoublem`toDoubler@.bigendian f.
fromWORD0=: fromWORDm`fromWORDr@.bigendian f.
fromDWORD0=: fromDWORDm`fromDWORDr@.bigendian f.
fromucode0=: fromucodem`fromucoder@.bigendian f.
fromDouble0=: fromDoublem`fromDoubler@.bigendian f.
toWORD1=: toWORDm`toWORDr@.(-.bigendian) f.
toDWORD1=: toDWORDm`toDWORDr@.(-.bigendian) f.
toucode1=: toucodem`toucoder@.(-.bigendian) f.
toDouble1=: toDoublem`toDoubler@.(-.bigendian) f.
fromWORD1=: fromWORDm`fromWORDr@.(-.bigendian) f.
fromDWORD1=: fromDWORDm`fromDWORDr@.(-.bigendian) f.
fromucode1=: fromucodem`fromucoder@.(-.bigendian) f.
fromDouble1=: fromDoublem`fromDoubler@.(-.bigendian) f.
DFH=: 3 : 0
if. '0x'-:2{.y=. }:^:('L'={:y) y do.
  d=. 0
  for_nib. ('0123456789abcdef'&i.) tolower 2}.y do.
    d=. nib (23 b.) 4 (33 b.) d
  end.
else.
  0&". y
end.
)
RGB=: 3 : 0"1
(0{y) (23 b.) 8 (33 b.) (1{y) (23 b.) 8 (33 b.) (2{y)
)

RGBtuple=: 3 : 0"0
(16bff (17 b.) y), (_8 (33 b.) 16bff00 (17 b.) y), (_16 (33 b.) 16bff0000 (17 b.) y)
)
3 : 0''
if. 504 < 0&". 'j'-.~4{.9!:14 '' do.
  fboxname=: ([: < 8 u: >) :: ]
else.
  fboxname=: ([: < [: (1&u: ::]) >) ::]
end.
empty ''
)

fread=: (1!:1 :: _1:) @ fboxname
fdir=: 1!:0@fboxname
freadx=: (1!:11 :: _1:)@(fboxname@{., }.)`(1!:11)@.(0: = L.)
fwritex=: ([ (1!:12) (fboxname@{., }.)@])`(1!:12)@.((0: = L.)@])
fopen=: (1!:21 :: _1:) @ (fboxname &>) @ boxopen
fclose=: (1!:22 :: _1:) @ (fboxname &>) @ boxopen
fwrite=: [ (1!:2) fboxname@]
fappend=: [ (1!:3) fboxname@]
fexist=: (1:@(1!:4) :: 0:) @ (fboxname &>) @ boxopen
ferase=: (1!:55 :: _1:) @ (fboxname &>) @ boxopen
maxpp=: 15 [ 16   
coclass 'oleheaderinfo'
coinsert 'olepps'
create=: 3 : 0
smallsize=: 16b1000
ppssize=: 16b80
bigblocksize=: 16b200
smallblocksize=: 16b0040
bdbcount=: 0
rootstart=: 0
sbdstart=: 0
sbdcount=: 0
extrabbdstart=: 0
extrabbdcount=: 0
bbdinfo=: 0 2$''
sbstart=: 0
sbsize=: 0
data=: ''
fileh=: ''
ppsfile=: ''
)

coclass 'olestorage'
coinsert 'oleutlfcn'
ppstyperoot=: 5
ppstypedir=: 1
ppstypefile=: 2
datasizesmall=: 16b1000
longintsize=: 4
ppssize=: 16b80
create=: 3 : 0
sfile=: y
openfilenum=: ''
headerinfo=: ''
)

destroy=: 3 : 0
if. '' -.@-: openfilenum do. fclose("0) openfilenum end.
if. '' -.@-: headerinfo do. destroy__headerinfo '' end.
codestroy ''
)
getppstree=: 3 : 0
bdata=. y
rhinfo=. initparse sfile
if. ''-:rhinfo do. '' return. end.
1{:: ugetppstree 0 ; rhinfo ; bdata ;< 0$0
)
getppssearch=: 3 : 0
'raname bdata icase'=. y
rhinfo=. initparse sfile
if. ''-:rhinfo do. '' return. end.
1{:: ugetppssearch 0 ; rhinfo ; raname ; bdata ; icase ;< 0$0
)
getnthpps=: 3 : 0
'ino bdata'=. y
rhinfo=. initparse sfile
if. ''-:rhinfo do. '' return. end.
1{:: ugetnthpps ino ; rhinfo ; <bdata
)
initparse=: 3 : 0
if. '' -.@-: headerinfo do. headerinfo return. end.
if. 1 4 e.~ 3!:0 y do.
  oio=. y
else.
  openfilenum=: ~. openfilenum, oio=. fopen <y   
end.
if. '' -.@-: p=. getheaderinfo oio do. headerinfo=: p end.
p
)
ugetppstree=: 3 : 0
'ino rhinfo bdata radone'=. y
if. #radone do.
  if. ino e. radone do. radone ; <'' return. end.
end.
irootblock=. rootstart__rhinfo
if. ''-: opps=. ugetnthpps ino ; rhinfo ; <bdata do. radone ; <'' return. end.
radone=. radone, ino
if. dirpps__opps ~: _1 do.
  'rad achild'=. ugetppstree dirpps__opps ; rhinfo ; bdata ; <radone
  radone=. radone, rad
  child__opps=: child__opps, achild
else.
  child__opps=: ''
end.
alist=. ''
if. prevpps__opps ~: _1 do.
  'rad achild'=. ugetppstree prevpps__opps ; rhinfo ; bdata ; <radone
  radone=. radone, rad
  alist=. alist, achild
end.
alist=. alist, opps
if. nextpps__opps ~: _1 do.
  'rad achild'=. ugetppstree nextpps__opps ; rhinfo ; bdata ; <radone
  radone=. radone, rad
  alist=. alist, achild
end.
radone ; <alist
)
ugetppssearch=: 3 : 0
'ino rhinfo raname bdata icase radone'=. y
irootblock=. rootstart__rhinfo
if. #radone do.
  if. ino e. radone do. radone ; <'' return. end.
end.
alist=. ''
if. ''-: opps=. ugetnthpps ino ; rhinfo ; <0 do. radone ; <'' return. end.
found=. 0
if. ((icase *. name__opps -:&toupper raname) +. name__opps-:raname) do.
  if. 1=bdata do.
    if. ''-: opps1=. ugetnthpps ino ; rhinfo ; <bdata do.
      destroy__opps ''
      radone ; <'' return.
    else.
      destroy__opps ''
      alist=. opps=. opps1
      radone=. radone, ino
    end.
  else.
    alist=. opps
    radone=. radone, ino
  end.
  found=. 1
end.
if. dirpps__opps ~: _1 do.
  'rad achild'=. ugetppssearch dirpps__opps ; rhinfo ; raname ; bdata ; icase ; <radone
  radone=. radone, rad
  alist=. alist, achild
end.
if. prevpps__opps ~: _1 do.
  'rad achild'=. ugetppssearch prevpps__opps ; rhinfo ; raname ; bdata ; icase ; <radone
  radone=. radone, rad
  alist=. alist, achild
end.
if. nextpps__opps ~: _1 do.
  'rad achild'=. ugetppssearch nextpps__opps ; rhinfo ; raname ; bdata ; icase ; <radone
  radone=. radone, rad
  alist=. alist, achild
end.
if. 0=found do. destroy__opps '' end.
radone ; <alist
)
getheaderinfo=: 3 : 0
fp=. 0
if. -. (freadx y, fp, 8)-:16bd0 16bcf 16b11 16be0 16ba1 16bb1 16b1a 16be1{a. do. '' return. end.
rhinfo=. '' conew 'oleheaderinfo'
fileh__rhinfo=: y
if. ''-:iwk=. getinfofromfile fileh__rhinfo ; 16b1e ; 2 do. '' [ destroy__rhinfo '' return. end.
bigblocksize__rhinfo=: <. 2&^ iwk
if. ''-:iwk=. getinfofromfile fileh__rhinfo ; 16b20 ; 2 do. '' [ destroy__rhinfo '' return. end.
smallblocksize__rhinfo=: <. 2&^ iwk
if. ''-:iwk=. getinfofromfile fileh__rhinfo ; 16b2c ; 4 do. '' [ destroy__rhinfo '' return. end.
bdbcount__rhinfo=: iwk
if. ''-:iwk=. getinfofromfile fileh__rhinfo ; 16b30 ; 4 do. '' [ destroy__rhinfo '' return. end.
rootstart__rhinfo=: iwk
if. ''-:iwk=. getinfofromfile fileh__rhinfo ; 16b3c ; 4 do. '' [ destroy__rhinfo '' return. end.
sbdstart__rhinfo=: iwk
if. ''-:iwk=. getinfofromfile fileh__rhinfo ; 16b40 ; 4 do. '' [ destroy__rhinfo '' return. end.
sbdcount__rhinfo=: iwk
if. ''-:iwk=. getinfofromfile fileh__rhinfo ; 16b44 ; 4 do. '' [ destroy__rhinfo '' return. end.
extrabbdstart__rhinfo=: iwk
if. ''-:iwk=. getinfofromfile fileh__rhinfo ; 16b48 ; 4 do. '' [ destroy__rhinfo '' return. end.
extrabbdcount__rhinfo=: iwk
bbdinfo__rhinfo=: getbbdinfo rhinfo
if. ''-: opps=. ugetnthpps 0 ; rhinfo ; <0 do. '' [ destroy__rhinfo '' return. end.
sbstart__rhinfo=: startblock__opps
sbsize__rhinfo=: size__opps
destroy__opps ''
rhinfo
)
getinfofromfile=: 3 : 0
'file ipos ilen'=. y
if. ''-:file do. '' return. end.
if. 2=ilen do.
  fromWORD0 freadx file, ipos, ilen
else.
  fromDWORD0 freadx file, ipos, ilen
end.
)
getbbdinfo=: 3 : 0
rhinfo=. y
abdlist=. ''
ibdbcnt=. bdbcount__rhinfo
i1stcnt=. <.(bigblocksize__rhinfo - 16b4c) % longintsize
ibdlcnt=. (<.bigblocksize__rhinfo % longintsize) - 1
fp=. 16b4c
igetcnt=. ibdbcnt <. i1stcnt
abdlist=. abdlist, fromDWORD0 freadx fileh__rhinfo, fp, longintsize*igetcnt
ibdbcnt=. ibdbcnt - igetcnt
iblock=. extrabbdstart__rhinfo
while. ((ibdbcnt> 0) *. isnormalblock iblock) do.
  fp=. setfilepos iblock ; 0 ; <rhinfo
  igetcnt=. ibdbcnt <. ibdlcnt
  abdlist=. abdlist, fromDWORD0 freadx fileh__rhinfo, fp, longintsize*igetcnt
  ibdbcnt=. ibdbcnt - igetcnt
  iblock=. fromDWORD0 freadx fileh__rhinfo, (fp=. fp+longintsize*igetcnt), longintsize
end.
iblkno=. 0
ibdcnt=. <.bigblocksize__rhinfo % longintsize
hbd=. 0 2$''
for_ibdl. abdlist do.
  fp=. setfilepos ibdl ; 0 ; <rhinfo
  awk=. fromDWORD0 freadx fileh__rhinfo, fp, bigblocksize__rhinfo
  for_i. i.ibdcnt do.
    if. ((i{awk) ~: iblkno+1) do.
      hbd=. hbd, iblkno, i{awk
    end.
    iblkno=. >:iblkno
  end.
end.
hbd
)
ugetnthpps=: 3 : 0
'ipos rhinfo bdata'=. y
ippsstart=. rootstart__rhinfo
ibasecnt=. <.bigblocksize__rhinfo % ppssize
ippsblock=. <.ipos % ibasecnt
ippspos=. ipos |~ ibasecnt
iblock=. getnthblockno ippsstart ; ippsblock ; <rhinfo
if. ''-:iblock do. '' return. end.
fp=. setfilepos iblock ; (ppssize*ippspos) ; <rhinfo
inmsize=. fromWORD0 (16b40+i.2){swk=. freadx fileh__rhinfo, fp, ppssize
inmsize=. 0 >. inmsize - 2
snm=. inmsize{.swk
sname=. fromucode0 (i.inmsize){swk
itype=. 256|fromWORD0 (16b42+i.2){swk
'lppsprev lppsnext ldirpps'=. fromDWORD0 (16b44+i.12){swk
ratime1st=: ((itype = ppstyperoot) +. (itype = ppstypedir)) { 0, oledate2local (16b64+i.8){swk
ratime2nd=: ((itype = ppstyperoot) +. (itype = ppstypedir)) { 0, oledate2local (16b6c+i.8){swk
'istart isize'=. fromDWORD0 (16b74+i.8){swk
if. 1=bdata do.
  sdata=. getdata itype ; istart ; isize ; <rhinfo
  pps=. createpps ipos ; sname ; itype ; lppsprev ; lppsnext ; ldirpps ; ratime1st ; ratime2nd ; istart ; isize ; sdata
else.
  pps=. createpps ipos ; sname ; itype ; lppsprev ; lppsnext ; ldirpps ; ratime1st ; ratime2nd ; istart ; isize
end.
pps
)
setfilepos=: 3 : 0
'iblock ipos rhinfo'=. y
ipos + (iblock+1)*bigblocksize__rhinfo
)
getnthblockno=: 3 : 0
'istblock inth rhinfo'=. y
inext=. istblock
for_i. i.inth do.
  isv=. inext
  inext=. getnextblockno isv ; <rhinfo
  if. 0= isnormalblock inext do. '' return. end.
end.
inext
)
getdata=: 3 : 0
'itype iblock isize rhinfo'=. y
if. itype = ppstypefile do.
  if. isize < datasizesmall do.
    getsmalldata iblock ; isize ; <rhinfo
  else.
    getbigdata iblock ; isize ; <rhinfo
  end.
elseif. itype = ppstyperoot do.  
  getbigdata iblock ; isize ; <rhinfo
elseif. itype = ppstypedir do.  
  0
end.
)
getbigdata=: 3 : 0
'iblock isize rhinfo'=. y
if. 0= isnormalblock iblock do. '' return. end.
irest=. isize
sres=. ''
akeys=. /:~ {.("1) bbdinfo__rhinfo
while. irest > 0 do.
  ares=. (akeys>:iblock)#akeys
  inkey=. {.ares
  i=. inkey - iblock
  inext=. ({:("1) bbdinfo__rhinfo){~({.("1) bbdinfo__rhinfo)i.inkey
  fp=. setfilepos iblock ; 0 ; <rhinfo
  igetsize=. irest <. bigblocksize__rhinfo * (i+1)
  sres=. sres, freadx fileh__rhinfo, fp, igetsize
  irest=. irest-igetsize
  iblock=. inext
end.
sres
)
getnextblockno=: 3 : 0
'iblockno rhinfo'=. y
if. iblockno e. {.("1) bbdinfo__rhinfo do.
  ({:("1) bbdinfo__rhinfo){~({.("1) bbdinfo__rhinfo)i.iblockno
else.
  iblockno+1
end.
)
isnormalblock=: 3 : 0
y -.@e. _4 _3 _2 _1
)
getsmalldata=: 3 : 0
'ismblock isize rhinfo'=. y
irest=. isize
sres=. ''
while. irest > 0 do.
  fp=. setfilepossmall ismblock ; <rhinfo
  sres=. sres, freadx fileh__rhinfo, fp, irest <. smallblocksize__rhinfo
  irest=. irest - smallblocksize__rhinfo
  ismblock=. getnextsmallblockno ismblock ; <rhinfo
end.
sres
)
setfilepossmall=: 3 : 0
'ismblock rhinfo'=. y
ismstart=. sbstart__rhinfo
ibasecnt=. <.bigblocksize__rhinfo % smallblocksize__rhinfo
inth=. <.ismblock%ibasecnt
ipos=. ismblock |~ ibasecnt
iblk=. getnthblockno ismstart ; inth ; <rhinfo
setfilepos iblk ; (ipos * smallblocksize__rhinfo) ; <rhinfo
)
getnextsmallblockno=: 3 : 0
'ismblock rhinfo'=. y
ibasecnt=. <.bigblocksize__rhinfo % longintsize
inth=. <.ismblock%ibasecnt
ipos=. ismblock |~ ibasecnt
iblk=. getnthblockno sbdstart__rhinfo ; inth ; <rhinfo
fp=. setfilepos iblk ; (ipos * longintsize) ; <rhinfo
fromDWORD0 freadx fileh__rhinfo, fp, longintsize
)

createpps=: 3 : 0
'ipos sname itype lppsprev lppsnext ldirpps ratime1st ratime2nd istart isize sdata'=. 11{.y
select. {.itype
case. ppstyperoot do.
  p=. conew 'oleppsroot'
  create__p ratime1st ; ratime2nd ;< ''
case. ppstypedir do.
  p=. conew 'oleppsdir'
  create__p sname ; ratime1st ; ratime2nd ;< ''
case. ppstypefile do.
  p=. conew 'oleppsfile'
  create__p sname ; sdata ;< ''
case. do.
  assert. 0
end.
no__p=: ipos
name__p=: u: sname
type__p=: {.itype
prevpps__p=: lppsprev
nextpps__p=: lppsnext
dirpps__p=: ldirpps
time1st__p=: ratime1st
time2nd__p=: ratime2nd
startblock__p=: istart
size__p=: isize
data__p=: sdata
p
)

coclass 'olepps'
coinsert 'oleutlfcn'
ppstyperoot=: 5
ppstypedir=: 1
ppstypefile=: 2
datasizesmall=: 16b1000
longintsize=: 4
ppssize=: 16b80
child=: ''
adoptedchild=: ''
destroy=: 3 : 0
for_pps. child -. adoptedchild do. destroy__pps '' end.
codestroy ''
)

fputs=: 3 : 0
if. fileh-:'' do. data=: data, y else. fileh fappend~ y end.
)
datalen=: 3 : 0
if. '' -.@-: ppsfile do. fsize ppsfile return. end.
#data
)
makesmalldata=: 3 : 0
'alist rhinfo'=. y
sres=. ''
ismblk=. 0
for_opps. alist do.
  if. type__opps=ppstypefile do.
    if. size__opps <: 0 do. continue. end.
    if. size__opps < smallsize__rhinfo do.
      ismbcnt=. >. size__opps % smallblocksize__rhinfo
      for_i. i.ismbcnt-1 do.
        fputs__rhinfo toDWORD0 i+ismblk+1
      end.
      fputs__rhinfo toDWORD0 _2
      if. '' -.@-: ppsfile__opps do.
        sres=. sres, ]`(''"_)@.(_1&-:)@:fread ppsfile__opps
      else.
        sres=. sres, data__opps
      end.
      if. size__opps |~ smallblocksize__rhinfo do.
        sres=. sres, ({.a.) #~ smallblocksize__rhinfo ([ - |) size__opps
      end.
      startblock__opps=: ismblk
      ismblk=. ismblk + ismbcnt
    end.
  end.
end.
isbcnt=. <. bigblocksize__rhinfo % longintsize
if. ismblk |~ isbcnt do.
  fputs__rhinfo, (,:toDWORD0 _1) #~ isbcnt ([ - |) ismblk
end.
sres
)
saveppswk=: 3 : 0
rhinfo=. y
z=. toucode0 name
z=. z, ({.a.)#~ 64-2*#name                         
z=. z, toWORD0 2*1+#name                     
z=. z, toBYTE type                                 
z=. z, toBYTE 16b00 
z=. z, toDWORD0 prevpps 
z=. z, toDWORD0 nextpps 
z=. z, toDWORD0 dirpps  
z=. z, 0 9 2 0{a.                                  
z=. z, 0 0 0 0{a.                                  
z=. z, 16bc0 0 0 0{a.                              
z=. z, 0 0 0 16b46{a.                              
z=. z, 0 0 0 0{a.                                  
z=. z, localdate2ole time1st                       
z=. z, localdate2ole time2nd                       
z=. z, toDWORD0 startblock                   
z=. z, toDWORD0 size                         
z=. z, toDWORD0 0                            
fputs__rhinfo z
z
)

coclass 'oleppsdir'
coinsert 'olepps'
create=: 3 : 0
'sname ratime1st ratime2nd rachild'=. y
no=: 0
name=: u: sname
type=: ppstypedir
prevpps=: 0 ] _1
nextpps=: 0 ] _1
dirpps=: 0 ] _1
time1st=: ratime1st
time2nd=: ratime2nd
startblock=: 0
size=: 0
data=: ''
child=: rachild
adoptedchild=: rachild
fileh=: ''
ppsfile=: ''
)

coclass 'oleppsfile'
coinsert 'olepps'
create=: 3 : 0
'snm sdata sfile'=. y
no=: 0
name=: u: snm
type=: ppstypefile
prevpps=: 0 ] _1
nextpps=: 0 ] _1
dirpps=: 0 ] _1
time1st=: 0
time2nd=: 0
startblock=: 0
size=: 0
data=: >(''-:sfile) { sdata ; ''
child=: ''
fileh=: ''
ppsfile=: ''
if. '' -.@-: sfile do.
  if. 1 4 e.~ 3!:0 sfile do.
    ppsfile=: sfile
  elseif. do.
    fname=. sfile
    ferase <fname
    ppsfile=: fopen <fname
  end.
  if. #sdata do.
    ppsfile fappend~ sdata
  end.
end.
)

append=: 3 : 0
if. '' -.@-: ppsfile do.
  ppsfile fappend~ y
else.
  data=: data, y
end.
)

coclass 'oleppsroot'
coinsert 'olepps'
create=: 3 : 0
'ratime1st ratime2nd rachild'=. y
no=: 0
name=: u: 'Root Entry'
type=: ppstyperoot
prevpps=: 0 ] _1
nextpps=: 0 ] _1
dirpps=: 0 ] _1
time1st=: ratime1st
time2nd=: ratime2nd
startblock=: 0
size=: 0
data=: ''
child=: rachild
adoptedchild=: rachild
fileh=: ''
ppsfile=: ''
)
save=: 3 : 0
'sfile bnoas rhinfo'=. y
if. ''-:rhinfo do.
  rhinfo=. '' conew 'oleheaderinfo'
end.
bigblocksize__rhinfo=: <. 2&^ (0= bigblocksize__rhinfo) { (adjust2 bigblocksize__rhinfo), 9
smallblocksize__rhinfo=: <. 2&^ (0= smallblocksize__rhinfo) { (adjust2 smallblocksize__rhinfo), 6
smallsize__rhinfo=: 16b1000
ppssize__rhinfo=: 16b80
if. ''-:sfile do.
  fileh__rhinfo=: ''
elseif. 1 4 e.~ 3!:0 sfile do.
  fileh__rhinfo=: sfile
elseif. do.
  ferase <sfile
  fileh__rhinfo=: fopen <sfile
end.
iblk=. 0
alist=. ''
list=. 18!:5 ''
if. bnoas do.
  alist=. 0{::saveppssetpnt2 list ; alist ; <rhinfo
else.
  alist=. 0{::saveppssetpnt list ; alist ; <rhinfo
end.
'isbdcnt ibbcnt ippscnt'=. calcsize alist ; <rhinfo
saveheader rhinfo ; isbdcnt ; ibbcnt ; <ippscnt
ssmwk=. makesmalldata alist ; <rhinfo
data=: ssmwk  
ibblk=. isbdcnt
ibblk=. savebigdata ibblk ; alist ; <rhinfo
savepps alist ; <rhinfo
savebbd isbdcnt ; ibbcnt ; ippscnt ; <rhinfo
if. (''-.@-:sfile) *. -. 1 4 e.~ 3!:0 sfile do.
  fclose fileh__rhinfo
end.
if. ''-:sfile do.
  rc=. data__rhinfo
else.
  rc=. ''
end.
destroy__rhinfo ''
rc
)
calcsize=: 3 : 0
'ralist rhinfo'=. y
isbdcnt=. 0
ibbcnt=. 0
ippscnt=. 0
ismalllen=. 0
isbcnt=. 0
for_opps. ralist do.
  if. type__opps=ppstypefile do.
    size__opps=: datalen__opps ''  
    if. size__opps < smallsize__rhinfo do.
      isbcnt=. isbcnt + >.size__opps % smallblocksize__rhinfo
    else.
      ibbcnt=. ibbcnt + >.size__opps % bigblocksize__rhinfo
    end.
  end.
end.
ismalllen=. isbcnt * smallblocksize__rhinfo
islcnt=. <. bigblocksize__rhinfo % longintsize
isbdcnt=. >.isbcnt % islcnt
ibbcnt=. ibbcnt + >.ismalllen % bigblocksize__rhinfo
icnt=. #ralist
ibdcnt=. <.bigblocksize__rhinfo % ppssize
ippscnt=. >.icnt % ibdcnt
isbdcnt ; ibbcnt ; <ippscnt
)
adjust2=: 3 : 0
>. 2^.y
)
saveheader=: 3 : 0
'rhinfo isbdcnt ibbcnt ippscnt'=. y
iblcnt=. <.bigblocksize__rhinfo % longintsize
i1stbdl=. <.(bigblocksize__rhinfo - 16b4c) % longintsize
i1stbdmax=. (i1stbdl*iblcnt) - i1stbdl
ibdexl=. 0
iall=. ibbcnt + ippscnt + isbdcnt
iallw=. iall
ibdcntw=. >.iallw % iblcnt
ibdcnt=. >.(iall + ibdcntw) % iblcnt
if. ibdcnt > 109 do.
  iblcnt=. <:iblcnt  
  ibbleftover=. iall - i1stbdmax
  if. iall > i1stbdmax do.
    whilst. ibdcnt > >. ibbleftover % iblcnt do.
      ibdcnt=. >. ibbleftover % iblcnt
      ibdexl=. >. ibdcnt % iblcnt
      ibbleftover=. ibbleftover + ibdexl
    end.
  end.
  ibdcnt=. ibdcnt + i1stbdl
end.
z=. 16bd0 16bcf 16b11 16be0 16ba1 16bb1 16b1a 16be1{a.
z=. z, 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0{a.
z=. z, toWORD0 16b3b
z=. z, toWORD0 16b03
z=. z, toWORD0 _2
z=. z, toWORD0 9
z=. z, toWORD0 6
z=. z, toWORD0 0
z=. z, 0 0 0 0 0 0 0 0 {a.
z=. z, toDWORD0 ibdcnt
z=. z, toDWORD0 ibbcnt+isbdcnt 
z=. z, toDWORD0 0
z=. z, toDWORD0 16b1000
z=. z, toDWORD0 (0=isbdcnt){0 _2    
z=. z, toDWORD0 isbdcnt
if. ibdcnt < i1stbdl do.
  z=. z, toDWORD0 _2       
  z=. z, toDWORD0 0        
else.
  z=. z, toDWORD0 iall+ibdcnt
  z=. z, toDWORD0 ibdexl
end.
fputs__rhinfo z
i=. 0
while. (i<i1stbdl) *. (i < ibdcnt) do.
  fputs__rhinfo toDWORD0 iall+i
  i=. >:i
end.
if. i<i1stbdl do.
  fputs__rhinfo, (,:toDWORD0 _1) #~ i1stbdl-i
end.
)
savebigdata=: 3 : 0
'istblk ralist rhinfo'=. y
ires=. 0
for_opps. ralist do.
  if. type__opps ~: ppstypedir do.
    size__opps=: datalen__opps ''   
    if. ((size__opps >: smallsize__rhinfo) +. ((type__opps = ppstyperoot) *. 0~:#data__opps)) do.
      if. '' -.@-: ppsfile__opps do.
        ilen=. #sbuff=. ]`(''"_)@.(_1&-:)@:fread ppsfile__opps
        fputs__rhinfo sbuff
      else.
        fputs__rhinfo data__opps
      end.
      if. size__opps |~ bigblocksize__rhinfo do.
        fputs__rhinfo ({.a.) #~ bigblocksize__rhinfo ([ - |) size__opps
      end.
      startblock__opps=: istblk
      istblk=. istblk + >.size__opps % bigblocksize__rhinfo
    end.
  end.
end.
istblk
)
savepps=: 3 : 0
'ralist rhinfo'=. y
for_pps. ralist do. saveppswk__pps rhinfo end.
icnt=. #ralist
ibcnt=. <.bigblocksize__rhinfo % ppssize__rhinfo
if. (icnt |~ ibcnt) do.
  fputs__rhinfo ({.a.) #~ (ibcnt ([ - |) icnt) * ppssize__rhinfo
end.
>.icnt % ibcnt
)
saveppssetpnt2=: 3 : 0
'athis ralist rhinfo'=. y
if. ''-:athis do. ralist ; _1 return.
elseif. 1=#athis do.
  ralist=. ralist, l=. 0{athis
  no__l=: (#ralist) -1
  prevpps__l=: _1
  nextpps__l=: _1
  ralist=. 0{::ra=. saveppssetpnt2 child__l ; ralist ; <rhinfo
  dirpps__l=: _1{::ra
  ralist ; no__l return.
elseif. do.
  icnt=. #athis
  ipos=. 0 
  awk=. athis
  if. (#athis) > 2 do.
    aprev=. 1{awk
    anext=. 2}.awk
  else.
    aprev=. ''
    anext=. }.awk
  end.
  l=. ipos{athis
  ralist=. 0{::ra=. saveppssetpnt2 aprev ; ralist ; <rhinfo
  prevpps__l=: _1{::ra
  ralist= ralist, l
  no__l=: (#ralist) -1
  ralist=. 0{::ra=. saveppssetpnt2 anext ; ralist ; <rhinfo
  nextpps__l=: _1{::ra
  ralist=. 0{::ra=. saveppssetpnt2 child__l ; ralist ; <rhinfo
  dirpps__l=: _1{::ra
  ralist ; no__l return.
end.
)
saveppssetpnt2s=: 3 : 0
'athis ralist rhinfo'=. y
if. (0:=#) athis do. ralist ; _1 return.
elseif. (#athis) =1 do.
  ralist=. ralist, l=. 0{athis
  no__l=: (#ralist) -1
  prevpps__l=: _1
  nextpps__l=: _1
  ralist=. 0{::ra=. saveppssetpnt2 child__l ; ralist ; <rhinfo
  dirpps__l=: _1{::ra
  ralist ; no__l return.
elseif. do.
  icnt=. #athis
  ipos=. 0 
  ralist=. ralist, l=. ipos{athis
  no__l=: (#ralist) -1
  awk=. athis
  aprev=. ipos{.awk
  anext=. (1+ipos)}.awk
  ralist=. 0{::ra=. saveppssetpnt2 aprev ; ralist ; <rhinfo
  prevpps__l=: _1{::ra
  ralist=. 0{::ra=. saveppssetpnt2 anext ; ralist ; <rhinfo
  nextpps__l=: _1{::ra
  ralist=. 0{::ra=. saveppssetpnt2 child__l ; ralist ; <rhinfo
  dirpps__l=: _1{::ra
  ralist ; no__l return.
end.
)
saveppssetpnt=: 3 : 0
'athis ralist rhinfo'=. y
if. (0:=#) athis do. ralist ; _1 return.
elseif. (#athis) =1 do.
  ralist=. ralist, l=. 0{athis
  no__l=: (#ralist) -1
  prevpps__l=: _1
  nextpps__l=: _1
  ralist=. 0{::ra=. saveppssetpnt child__l ; ralist ; <rhinfo
  dirpps__l=: _1{::ra
  ralist ; no__l return.
elseif. do.
  icnt=. #athis
  ipos=. <.icnt % 2
  ralist=. ralist, l=. ipos{athis
  no__l=: (#ralist) -1
  awk=. athis
  aprev=. ipos{.awk
  anext=. (1+ipos)}.awk
  ralist=. 0{::ra=. saveppssetpnt aprev ; ralist ; <rhinfo
  prevpps__l=: _1{::ra
  ralist=. 0{::ra=. saveppssetpnt anext ; ralist ; <rhinfo
  nextpps__l=: _1{::ra
  ralist=. 0{::ra=. saveppssetpnt child__l ; ralist ; <rhinfo
  dirpps__l=: _1{::ra
  ralist ; no__l return.
end.
)
saveppssetpnt1=: 3 : 0
'athis ralist rhinfo'=. y
if. (0:=#) athis do. ralist ; _1 return.
elseif. (#athis) =1 do.
  ralist=. ralist, l=. 0{athis
  no__l=: (#ralist) -1
  prevpps__l=: _1
  nextpps__l=: _1
  ralist=. 0{::ra=. saveppssetpnt child__l ; ralist ; <rhinfo
  dirpps__l=: _1{::ra
  ralist ; no__l return.
elseif. do.
  icnt=. #athis
  ipos=. <.icnt % 2
  ralist=. ralist, l=. ipos{athis
  no__l=: (#ralist) -1
  awk=. athis
  aprev=. ipos{.awk
  anext=. (1+ipos)}.awk
  ralist=. 0{::ra=. saveppssetpnt aprev ; ralist ; <rhinfo
  prevpps__l=: _1{::ra
  ralist=. 0{::ra=. saveppssetpnt anext ; ralist ; <rhinfo
  nextpps__l=: _1{::ra
  ralist=. 0{::ra=. saveppssetpnt child__l ; ralist ; <rhinfo
  dirpps__l=: _1{::ra
  ralist ; no__l return.
end.
)
savebbd=: 3 : 0
'isbdsize ibsize ippscnt rhinfo'=. y
ibbcnt=. <.bigblocksize__rhinfo % longintsize
i1stbdl=. <.(bigblocksize__rhinfo - 16b4c) % longintsize
ibdexl=. 0
iall=. ibsize + ippscnt + isbdsize
iallw=. iall
ibdcntw=. >.iallw % ibbcnt
ibdcnt=. >.(iall + ibdcntw) % ibbcnt
if. ibdcnt >i1stbdl do.
  whilst. ibdcnt > i1stbdl+ibdexl*ibbcnt do.
    ibdexl=. >:ibdexl
    iallw=. >:iallw
    ibdcntw=. >.iallw % ibbcnt
    ibdcnt=. >.(iallw + ibdcntw) % ibbcnt
  end.
end.
if. isbdsize > 0 do.
  for_i. i.(isbdsize-1) do.
    fputs__rhinfo toDWORD0 i+1
  end.
  fputs__rhinfo toDWORD0 _2
end.
for_i. i.(ibsize-1) do.
  fputs__rhinfo toDWORD0 i+isbdsize+1
end.
fputs__rhinfo toDWORD0 _2
for_i. i.(ippscnt-1) do.
  fputs__rhinfo toDWORD0 i+isbdsize+ibsize+1
end.
fputs__rhinfo toDWORD0 _2
for_i. i.ibdcnt do.
  fputs__rhinfo toDWORD0 _3
end.
for_i. i.ibdexl do.
  fputs__rhinfo toDWORD0 _4
end.
if. ((iallw + ibdcnt) |~ ibbcnt) do.
  fputs__rhinfo, (,:toDWORD0 _1) #~ ibbcnt ([ - |) (iallw + ibdcnt)
end.
if. (ibdcnt > i1stbdl) do.
  in=. 0
  inb=. 0
  i=. i1stbdl
  while. i<ibdcnt do.
    if. (in>: (ibbcnt-1)) do.
      in=. 0
      inb=. >:inb
      fputs__rhinfo toDWORD0 iall+ibdcnt+inb
    end.
    fputs__rhinfo toDWORD0 ibsize+isbdsize+ippscnt+i
    i=. >:i
    in=. >:in
  end.
  if. ((ibdcnt-i1stbdl) |~ (ibbcnt-1)) do.
    fputs__rhinfo, (,:toDWORD0 _1) #~ (ibbcnt-1) ([ - |) (ibdcnt-i1stbdl)
  end.
  fputs__rhinfo toDWORD0 _2
end.
)
IFWINE_z_=: (6 = 9!:12'') *. 0-.@-:2!:5 'LOGNAME'
coclass 'biff'
coinsert 'oleutlfcn'
shortdatefmt=: 'dd/mm/yyyy'
RECORDLEN=: 8224   
format0n=: 164  
colorset0n=: 8   
colorborder=: 16b40
colorpattern=: 16b41
colorfont=: 16b7fff
DEBUG=: 1
A1toRC=: 3 : 0
assert. y e. '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ'
(<: 26 #. ' ABCDEFGHIJKLMNOPQRSTUVWXYZ'&i. _5&{. c),~ <: 1&". y -. c=. y -. '0123456789'
)

toHeader=: toWORD0
UString=: 3 : 0
toucode0 u: y
)

SString=: 3 : 0
(1&u: ::]) y
)

toString=: 3 : 0
if. 131072= 3!:0 y do.
  toucode0 y
else.
  y
end.
)

toUString8=: 3 : 0
if. 131072= 3!:0 y do.
  (a.{~#y), (1{a.), toucode0 y
else.
  (a.{~#y), (0{a.), y
end.
)

toUString16=: 3 : 0
if. 131072= 3!:0 y do.
  (toWORD0 #y), (1{a.), toucode0 y
else.
  (toWORD0 #y), (0{a.), y
end.
)

toUString=: 3 : 0
if. 131072= 3!:0 y do.
  (1{a.), toucode0 y
else.
  (0{a.), y
end.
)

toStream=: 4 : 0
x fappend~ y
)

sulen8=: 3 : 0
if. 131072= 3!:0 y do.
  2+2*#y
else.
  2+#y
end.
)

sulen16=: 3 : 0
if. 131072= 3!:0 y do.
  3+2*#y
else.
  3+#y
end.
)
cellborder_no_line=: 0
cellborder_thin=: 1
cellborder_medium=: 2
cellborder_dashed=: 3
cellborder_dotted=: 4
cellborder_thick=: 5
cellborder_double=: 6
cellborder_hair=: 7
cellborder_medium_dashed=: 8
cellborder_thin_dash_dotted=: 9
cellborder_medium_dash_dotted=: 10
cellborder_thin_dash_dot_dotted=: 11
cellborder_medium_dash_dot_dotted=: 12
cellborder_slanted_medium_dash_dotted=: 13
coinstance=: {. @ (18!:2 & boxxopen)
biffappend=: 3 : 0

0 biffappend y
:

if. RECORDLEN < #y do. x add_continue y else. y end.
)
add_continue=: 4 : 0

data=. y
recordtype=. 16b003c 
if. x do. data return. end.
data=. RECORDLEN}.data [ tmp=. RECORDLEN{.data
tmp=. (toWORD0 RECORDLEN-4) 2 3}tmp
while. RECORDLEN < #data do.
  tmp=. tmp, (RECORDLEN{.data),~ toHeader recordtype, RECORDLEN
  data=. RECORDLEN}.data
end.
tmp=. tmp, data,~ toHeader recordtype, #data
)
biffcrc32=: 3 : 0

2&(3!:4) 4# (128!:3) y
)

biffmd4=: 3 : 0

md5bin_crypt_ y
)
add_mso_generic=: 4 : 0

data=. x
'type version instance length'=. y
header=. version (23 b.) 4 (33 b.) instance

record=. toWORD0 header, type
record=. record, toDWORD0 length
record, data
)
biff_array=: 3 : 0
'range recalc parsedexpr'=. y
'firstrow lastrow firstcol lastcol'=. range
recordtype=. 16b0221
z=. ''
z=. z, toWORD0 firstrow
z=. z, toWORD0 lastrow
z=. z, toBYTE firstcol
z=. z, toBYTE lastcol
z=. z, toWORD0 recalc
z=. z, toDWORD0 0
z=. z, toWORD0 #parsedexpr
z=. z, SString parsedexpr
z=. z,~ toHeader recordtype, #z
)
biff_backup=: 3 : 0
recordtype=. 16b0040
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)
biff_blank=: 4 : 0
recordtype=. 16b0201
y=. >y
assert. 2=#y
z=. ''
z=. z, toWORD0 y
z=. z, toWORD0 x
z=. z,~ toHeader recordtype, #z
)
biff_bof=: 3 : 0
'version docn'=. y
recordtype=. 16b809
z=. ''
z=. z, toWORD0 version
z=. z, toWORD0 docn
z=. z, toWORD0 8111   
z=. z, toWORD0 1997   
z=. z, toDWORD0 16b40c1
z=. z, toDWORD0 16b106
z=. z,~ toHeader recordtype, #z
)

biff_bookbool=: 3 : 0
recordtype=. 16b00da
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)
biff_boolerr=: 4 : 0
'rowcol boolerrvalue boolORerr'=. y
recordtype=. 16b0205
z=. ''
z=. z, toWORD0 rowcol
z=. z, toWORD0 x
z=. z, toBYTE boolerrvalue
z=. z, toBYTE boolORerr
z=. z,~ toHeader recordtype, #z
)
biff_bottommargin=: 3 : 0
recordtype=. 16b0029
z=. ''
z=. z, toDouble0 y
z=. z,~ toHeader recordtype, #z
)

biff_boundsheet=: 3 : 0
'offset visible type sheetname'=. y
recordtype=. 16b85
z=. ''
z=. z, toDWORD0 offset
z=. z, toBYTE visible
z=. z, toBYTE type
z=. z, toUString8 sheetname
z=. z,~ toHeader recordtype, #z
)
biff_calccount=: 3 : 0
recordtype=. 16b000c
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)
biff_calcmode=: 3 : 0
recordtype=. 13
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)

biff_codepage=: 3 : 0
recordtype=. 16b0042
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)

biff_colinfo=: 4 : 0
'col1 col2 width hide level collapse'=. y
recordtype=. 16b007d
z=. ''
z=. z, toWORD0 col1, col2, width, x
z=. z, toWORD0 hide bitor 8 bitshl level bitor 4 bitshl collapse
z=. z, toWORD0 0
z=. z,~ toHeader recordtype, #z
)
biff_columndefault=: 3 : 0
recordtype=. 16b0020
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)
biff_continue=: 3 : 0
recordtype=. 16b003c
z=. ''
z=. z, toString y
z=. z,~ toHeader recordtype, #z
)

biff_country=: 3 : 0
recordtype=. 16b008c
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)

biff_crn=: 3 : 0
recordtype=. 16b005a
z=. ''
z=. z, SString y
z=. z,~ toHeader recordtype, #z
)
biff_date1904=: 3 : 0
recordtype=. 16b0022
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)
biff_defaultcolwidth=: 3 : 0
recordtype=. 16b0055
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)
biff_defaultrowheight=: 3 : 0
'notmatch hidden spaceabove spacebelow height'=. y
recordtype=. 16b0225
z=. ''
z=. z, toWORD0 bitor (0~:notmatch) bitor 1 bitshl (0~:hidden) bitor 1 bitshl (0~:spaceabove) bitor 1 bitshl (0~:spacebelow)
z=. z, toWORD0 height
z=. z,~ toHeader recordtype, #z
)
biff_delta=: 3 : 0
recordtype=. 16b0010
z=. ''
z=. z, toDouble0 y
z=. z,~ toHeader recordtype, #z
)
biff_dimensions=: 3 : 0
recordtype=. 16b0200
z=. ''
z=. z, toDWORD0 0 1+0 1{y
z=. z, toWORD0 0 1+2 3{y
z=. z, toWORD0 0
z=. z,~ toHeader recordtype, #z
)
biff_eof=: 3 : 0
recordtype=. 16b000a
z=. ''
z=. z,~ toHeader recordtype, #z
)

biff_externname=: 4 : 0
recordtype=. 16b0023
z=. ''
if. 'external'-:x do.
  'builtin sheetidx name formula'
  z=. z, toWORD0 0~:builtin
  z=. z, toWORD0 >:sheetidx
  z=. z, toWORD0 0
  z=. z, toUString8 name
  z=. z, SString formula
elseif. 'internal'-:x do.
  'unhandled exception' 13!:8 (3)
elseif. 'addin'-:x do.
  z=. z, toWORD0 0
  z=. z, toDWORD0 0
  z=. z, toUString8 y
  z=. z, SString a.{~2 0 16b1c 16b17
elseif. 'dde'-:x do.
  'automatic stddoc clip item'=. y
  if. 0~:stddoc do. clip=. 16bfff end.
  z=. z, toWORD0 1 bitshl (0~:automatic) bitand 1 bitshl (0~:stddoc) bitand 2 bitshl clip
  z=. z, toDWORD0 0
  z=. z, toUString8 item
elseif. 'ole'-:x do.
  'automatic storage'=. y
  z=. z, toWORD0 1 bitshl (0~:automatic)
  z=. z, toDWORD0 storage
  z=. z, SString a.{~1 0 16b27
elseif. do.
  'unhandled exception' 13!:8 (3)
end.
z=. z,~ toHeader recordtype, #z
)

biff_externsheet=: 3 : 0
recordtype=. 16b0017
z=. ''
z=. z, toWORD0 #y
for_yi. y do.
  z=. z, toWORD0 yi
end.
z=. z,~ toHeader recordtype, #z
)
biff_font=: 3 : 0
recordtype=. 16b0031
z=. ''
'height italic strikeout color weight script underline family charset fontname'=. y
z=. z, toWORD0 height
z=. z, toWORD0 (1 bitshl 0~:italic) bitor (3 bitshl 0~:strikeout)
z=. z, toWORD0 color
z=. z, toWORD0 weight
z=. z, toWORD0 script
z=. z, toBYTE underline
z=. z, toBYTE family
z=. z, toBYTE charset
z=. z, toBYTE 0
z=. z, toUString8 fontname
z=. z,~ toHeader recordtype, #z
)
biff_footer=: 3 : 0
recordtype=. 16b0015
z=. ''
if. #y do.
  z=. z, toUString16 y
end.
z=. z,~ toHeader recordtype, #z
)
biff_format=: 3 : 0
'num str'=. y
recordtype=. 16b041e
z=. ''
z=. z, toWORD0 num
z=. z, toUString16 str
z=. z,~ toHeader recordtype, #z
)
biff_formula=: 4 : 0
'rowcol value recalc calcopen shared parsedexpr'=. y
recordtype=. 16b0006
z=. ''
z=. z, toWORD0 rowcol
z=. z, toWORD0 x
z=. z, SString 8{.value
z=. z, toWORD0 (0~:recalc) bitor 1 bitshl 1 (0~:calcopen) bitor 1 bitshl 1 (0~:shared)
z=. z, toDWORD0 0
z=. z, toWORD0 #parsedexpr
z=. z, SString parsedexpr
z=. z,~ toHeader recordtype, #z
)

biff_gridset=: 3 : 0
recordtype=. 16b0082
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)

biff_guts=: 3 : 0
recordtype=. 16b0080
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)

biff_hcenter=: 3 : 0
recordtype=. 16b0083
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)
biff_header=: 3 : 0
recordtype=. 16b0014
z=. ''
if. #y do.
  z=. z, toUString16 y
end.
z=. z,~ toHeader recordtype, #z
)

biff_hideobj=: 3 : 0
recordtype=. 16b008d
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)

biff_hlink=: 3 : 0
recordtype=. 16b01b8
z=. ''
'linktype rowcols link elink uplevel textmark target description'=. y
bit8=. bit7=. bit6=. bit5=. bit4=. bit3=. bit2=. bit1=. bit0=. 0
if. #target do. bit7=. 1 end.
if. #description do. bit2=. bit4=. 1 end.
if. #textmark do. bit3=. 1
elseif. 'worksheet'-:linktype do.
  'unhandled exception' 13!:8 (3)
end.
if. ('worksheet'-:linktype) *. (':'e.link) do. bit1=. 1 end.
if. 'url'-:linktype do. bit0=. bit1=. 1
elseif. 'local'-:linktype do. bit0=. 1
elseif. 'unc'-:linktype do. bit0=. bit1=. bit8=. 1
elseif. 'worksheet'-:linktype do. bit3=. 1
end.
flag=. #. bit8, bit7, bit6, bit5, bit4, bit3, bit2, bit1, bit0
z=. z, toWORD0 rowcols  
z=. z, SString a.{~16bd0 16bc9 16bea 16b79 16bf9 16bba 16bce 16b11 16b8c 16b82 16b00 16baa 16b00 16b4b 16ba9 16b0b
z=. z, toDWORD0 2
z=. z, toDWORD0 flag
if. #description do.
  z=. z, toDWORD0 1+#description
  z=. z, UString description,{.a.
end.
if. #target do.
  z=. z, toDWORD0 1+#target
  z=. z, UString target,{.a.
end.
if. 'url'-:linktype do.
  z=. z, SString a.{~16be0 16bc9 16bea 16b79 16bf9 16bba 16bce 16b11 16b8c 16b82 16b00 16baa 16b00 16b4b 16ba9 16b0b
  z=. z, toDWORD0 2*1+#link
  z=. z, UString link,{.a.
elseif. 'local'-:linktype do.
  z=. z, SString a.{~16b03 16b03 16b00 16b00 16b00 16b00 16b00 16b00 16bc0 16b00 16b00 16b00 16b00 16b00 16b00 16b46
  z=. z, toWORD0 uplevel
  z=. z, toDWORD0 1+#link
  z=. z, SString link,{.a.
  z=. z, SString a.{~16bff 16bff 16bad 16bde, 20#0
  if. #elink do.
    z=. z, toDWORD0 10+2*#elink
    z=. z, toDWORD0 2*#elink
    z=. z, SString a.{~16b03 16b00
    z=. z, UString elink
  else.
    z=. z, toDWORD0 0
  end.
elseif. 'unc'-:linktype do.
  z=. z, SString a.{~16be0 16bc9 16bea 16b79 16bf9 16bba 16bce 16b11 16b8c 16b82 16b00 16baa 16b00 16b4b 16ba9 16b0b
  z=. z, toDWORD0 1+#link
  z=. z, UString link,{.a.
elseif. 'worksheet'-:linktype do.
  ''  
end.
if. #textmark do.
  z=. z, toDWORD0 1+#textmark
  z=. z, UString textmark,{.a.
end.
z=. z,~ toHeader recordtype, #z
)
biff_horizontalpagebreaks=: 3 : 0
recordtype=. 16b001b
z=. ''
z=. z, toWORD0 #y
for_iy. y do. z=. z, toWORD0 iy end.
z=. z,~ toHeader recordtype, #z
)

biff_index=: 3 : 0
recordtype=. 16b020b
z=. ''
z=. z, toDWORD0 0
z=. z, toDWORD0 y
z=. z,~ toHeader recordtype, #z
)
biff_integer=: 4 : 0
assert. 2=#y
assert. 2=# 0{::y
assert. 1 4 8 e.~ (3!:0) 1{::y
recordtype=. 16b027e
z=. ''
z=. z, toWORD0 0{::y
z=. z, toWORD0 x
z=. z, toDWORD0 2b10 bitor 2 bitshl <. 1{::y
z=. z,~ toHeader recordtype, #z
)
biff_iteration=: 3 : 0
recordtype=. 16b0011
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)
biff_label=: 4 : 0
assert. 2=#y
assert. 2=# 0{::y
assert. 2 131072 e.~ (3!:0) 1{::y
if. ''-:, 1{::y do.
  x biff_blank {.y
else.
  recordtype=. 16b00fd
  z=. ''
  z=. z, toWORD0 0{::y
  z=. z, toWORD0 x
  z=. z, toDWORD0 add2sst ,&.> 1{y
  z=. z,~ toHeader recordtype, #z
end.
)

biff_labelranges=: 3 : 0
'row col'=. y
if. 0=(#row)+#col do. '' return. end.
recordtype=. 16b015f
z=. ''
z=. z, toWORD0 #row
for_iy. row do. z=. z, toWORD0 iy end.
z=. z, toWORD0 #col
for_iy. col do. z=. z, toWORD0 iy end.
z=. z,~ toHeader recordtype, #z
)
biff_leftmargin=: 3 : 0
recordtype=. 16b0026
z=. ''
z=. z, toDouble0 y
z=. z,~ toHeader recordtype, #z
)

biff_mergedcells=: 3 : 0
if. (0:=#) y do. '' return. end.
recordtype=. 16b00e5
z=. ''
z=. z, toWORD0 #y
for_iy. y do. z=. z, toWORD0 iy end.
z=. z,~ toHeader recordtype, #z
)

biff_name=: 3 : 0
recordtype=. 16b0018
z=. ''
'hidden function command macro complex builtin functiongroup binaryname keybd name formula sheet menu description helptopic statusbar'=. y
flag=. (0~:hidden) bitor 1 bitshl (0~:function) bitor 1 bitshl (0~:command) bitor 1 bitshl (0~:macro) bitor 1 bitshl (0~:complex) bitor 1 bitshl (0~:builtin) bitor 1 bitshl functiongroup bitor 6 bitshl binaryname
z=. z, toWORD0 flag
z=. z, toBYTE keybd
z=. z, toBYTE #name
z=. z, toWORD0 #formula
z=. z, toWORD0 0
z=. z, toWORD0 >:sheet
z=. z, toBYTE #menu
z=. z, toBYTE #description
z=. z, toBYTE #helptopic
z=. z, toBYTE #statusbar
z=. z, toUString name
z=. z, SString formula
if. #menu do. z=. z, toUString menu end.
if. #description do. z=. z, toUString description end.
if. #helptopic do. z=. z, toUString helptopic end.
if. #statusbar do. z=. z, toUString statusbar end.
z=. z,~ toHeader recordtype, #z
)
biff_number=: 4 : 0
assert. 2=#y
assert. 2=# 0{::y
assert. 1 4 8 e.~ (3!:0) 1{::y
recordtype=. 16b0203
z=. ''
z=. z, toWORD0 0{::y
z=. z, toWORD0 x
z=. z, toDouble0 1{::y
z=. z,~ toHeader recordtype, #z
)
biff_objectprotect=: 3 : 0
recordtype=. 16b0063
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)

biff_palette=: 3 : 0
recordtype=. 16b0092
z=. ''
z=. z, toWORD0 #y
z=. z, toDWORD0 y
z=. z,~ toHeader recordtype, #z
)
biff_pane=: 3 : 0
recordtype=. 16b0041
z=. ''
'split vis activepane'=. y
z=. z, toWORD0 split   
z=. z, toWORD0 vis     
for_pane. activepane do. z=. z, toWORD0 pane end.
z=. z,~ toHeader recordtype, #z
)
biff_password=: 3 : 0
recordtype=. 16b0013
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)
biff_precision=: 3 : 0
recordtype=. 16b000e
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)
biff_printgridlines=: 3 : 0
recordtype=. 16b002b
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)
biff_printheaders=: 3 : 0
recordtype=. 16b002a
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)
biff_protect=: 3 : 0
recordtype=. 16b0012
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)
biff_quicktip=: 3 : 0
recordtype=. 16b0800
z=. ''
'rowcols description'=. y
z=. z, toWORD0 rowcols  
z=. z, UString description,{.a.
z=. z,~ toHeader recordtype, #z
)
biff_refmode=: 3 : 0
recordtype=. 16b000f
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)
biff_rightmargin=: 3 : 0
recordtype=. 16b0027
z=. ''
z=. z, toDouble0 y
z=. z,~ toHeader recordtype, #z
)
biff_row=: 4 : 0
xf=. x
'rownumber firstcol lastcol usedefaultheight rowheight heightnotmatch spaceabove spacebelow hidden explicitformat outlinelevel outlinegroup'=. y
recordtype=. 16b0208
z=. ''
z=. z, toWORD0 rownumber
z=. z, toWORD0 firstcol
z=. z, toWORD0 lastcol
z=. z, toWORD0 (16b7fff bitand rowheight) bitor 15 bitshl 0~:usedefaultheight
z=. z, toWORD0 0
z=. z, toWORD0 0
z=. z, toDWORD0 (16b7 bitand outlinelevel) bitor 4 bitshl (0~:outlinegroup) bitor 1 bitshl (0~:hidden) bitor 1 bitshl (0~:heightnotmatch) bitor 1 bitshl (0~:explicitformat) bitor 1 bitshl 1 bitor 8 bitshl (16bfff bitand xf) bitor 12 bitshl (0~:spaceabove) bitor 1 bitshl 0~:spacebelow
z=. z,~ toHeader recordtype, #z
)
biff_scenprotect=: 3 : 0
recordtype=. 16b00dd
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)

biff_scl=: 3 : 0
recordtype=. 16b00a0
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)
biff_selection=: 3 : 0
recordtype=. 16b001d
z=. ''
'panenum row col refnum refs'=. y
z=. z, toBYTE panenum
z=. z, toWORD0 row
z=. z, toWORD0 col
z=. z, toWORD0 refnum
z=. z, toWORD0 #refs
for_i. i.#refs do.
  z=. z, toWORD0 2{.>i{refs
  z=. z, toBYTE 2}.>i{refs
end.
z=. z,~ toHeader recordtype, #z
)

biff_setup=: 3 : 0
recordtype=. 16b00a1
z=. ''
'setuppapersize setupscaling setupstartpage setupfitwidth setupfitheight setuprowmajor setupportrait setupinvalid setupblackwhite setupdraft setupcellnote setuporientinvalid setupusestartpage setupnoteatend setupprinterror setupdpi setupvdpi setupheadermargin setupfootermargin setupnumcopy'=. y
z=. z, toWORD0 setuppapersize
z=. z, toWORD0 setupscaling
z=. z, toWORD0 setupstartpage
z=. z, toWORD0 setupfitwidth
z=. z, toWORD0 setupfitheight
z=. z, toWORD0 setuprowmajor bitor 1 bitshl setupportrait bitor 1 bitshl setupinvalid bitor 1 bitshl setupblackwhite bitor 1 bitshl setupdraft bitor 1 bitshl setupcellnote bitor 1 bitshl setuporientinvalid bitor 1 bitshl setupusestartpage bitor 2 bitshl setupnoteatend bitor 1 bitshl setupprinterror
z=. z, toWORD0 setupdpi
z=. z, toWORD0 setupvdpi
z=. z, toDouble0 setupheadermargin
z=. z, toDouble0 setupfootermargin
z=. z, toWORD0 setupnumcopy
z=. z,~ toHeader recordtype, #z
)

biff_standardwidth=: 3 : 0
recordtype=. 16b0099
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)
biff_string=: 3 : 0
if. #y do.
  recordtype=. 16b0207
  z=. ''
  z=. z, toUString16 y
  z=. z,~ toHeader recordtype, #z
else.
  z=. ''
end.
)

biff_style=: 4 : 0
recordtype=. 16b0293
z=. ''
'builtin id level name'=. y
if. 0=builtin do.
  z=. z, toWORD0 16bfff bitand x
  z=. z, toUString16 name
else.
  z=. z, toWORD0 16b8000 bitor 16bfff bitand x
  z=. z, toBYTE id
  z=. z, toBYTE level
end.
z=. z,~ toHeader recordtype, #z
)

biff_supbook=: 4 : 0
recordtype=. 16b01ae
z=. ''
if. 'external'-:x do.
  z=. z, toWORD0 # _1{::y
  z=. z, toUString16 0{::y
  for_yi. _1{::y do.
    z=. z, toUString16 >yi
  end.
elseif. 'internal'-:x do.
  z=. z, toWORD0 y
  z=. z, SString 1 4{a.
elseif. 'addin'-:x do.
  z=. z, toWORD0 1
  z=. z, SString 1 3{a.
elseif. ('ole'-:x)+.('dde'-:x) do.
  z=. z, toWORD0 0
  z=. z, toUString16 y
elseif. do.
  'unhandled exception' 13!:8 (3)
end.
z=. z,~ toHeader recordtype, #z
)
biff_table=: 3 : 0
recordtype=. 16b0036
z=. ''
'firstrow lastrow firstcol lastcol recalc rowinput colinput row col'=. y
z=. z, toWORD0 firstrow
z=. z, toWORD0 lastrow
z=. z, toBYTE firstcol
z=. z, toBYTE lastcol
z=. z, toWORD0 recalc
z=. z, toWORD0 rowinput
z=. z, toWORD0 colinput
z=. z, toWORD0 row
z=. z, toWORD0 col
z=. z,~ toHeader recordtype, #z
)
biff_topmargin=: 3 : 0
recordtype=. 16b0028
z=. ''
z=. z, toDouble0 y
z=. z,~ toHeader recordtype, #z
)

biff_vcenter=: 3 : 0
recordtype=. 16b0084
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)
biff_verticalpagebreaks=: 3 : 0
recordtype=. 16b001a
z=. ''
z=. z, toWORD0 #y
for_iy. y do. z=. z, toWORD0 iy end.
z=. z,~ toHeader recordtype, #z
)
biff_window1=: 3 : 0
recordtype=. 16b003d
z=. ''
'x y width height hidden'=. y
z=. z, toWORD0 x
z=. z, toWORD0 y
z=. z, toWORD0 width
z=. z, toWORD0 height
z=. z, toWORD0 hidden
z=. z, toWORD0 0
z=. z, toWORD0 0
z=. z, toWORD0 1
z=. z, toWORD0 250
z=. z,~ toHeader recordtype, #z
)
biff_window2=: 3 : 0
recordtype=. 16b023e
z=. ''
'flag topvisiblerow leftvisiblecol RGBcolor'=. y
z=. z, toWORD0 flag
z=. z, toWORD0 topvisiblerow
z=. z, toWORD0 leftvisiblecol
z=. z, toWORD0 RGBcolor
z=. z, toWORD0 0
z=. z, toWORD0 0
z=. z, toWORD0 0
z=. z, toDWORD0 0
z=. z,~ toHeader recordtype, #z
)
biff_windowprotect=: 3 : 0
recordtype=. 16b0019
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)

biff_wsbool=: 3 : 0
recordtype=. 16b0081
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)

biff_xct=: 3 : 0
recordtype=. 16b0059
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)

biff_xf=: 3 : 0
recordtype=. 16b00e0
'font format typeprot align rotate indent used border linecolor color'=. y
z=. ''
z=. z, toWORD0 font
z=. z, toWORD0 format
z=. z, toWORD0 typeprot
z=. z, toBYTE align
z=. z, toBYTE rotate
z=. z, toBYTE indent
z=. z, toBYTE used
z=. z, toDWORD0 border
z=. z, toDWORD0 linecolor
z=. z, toWORD0 color
z=. z,~ toHeader recordtype, #z
)

coclass 'biffrefname'
coinsert 'biff'
create=: 3 : 0
'hidden function command macro complex builtin functiongroup binaryname keybd sheetidx'=: 0
'name formula menu description helptopic statusbar'=: 6#a:
)

destroy=: codestroy
writestream=: 3 : 0
z=. biff_name hidden ; function ; command ; macro ; complex ; builtin ; functiongroup ; binaryname ; keybd ; name ; formula ; sheetidx ; menu ; description ; helptopic ; statusbar
)

coclass 'biffsupbook'
coinsert 'biff'
newextname=: 3 : 0
extname=: extname, <y
)

create=: 3 : 0
'type sheetn source sheetname'=: 4{.y
if. -. (<type) e. 'external' ; 'internal' ; 'addin' ; 'ole' ; 'dde' do.
  'unhandled exception' 13!:8 (3)
end.
extname=: ''
crn=: ''  
)

destroy=: codestroy
writestream=: 3 : 0
z=. ''
if. 'external'-:type do. z=. z, type biff_supbook source ,&< sheetname
elseif. 'internal'-:type do. z=. z, type biff_supbook y
elseif. 'addin'-:type do. z=. z, type biff_supbook ''
elseif. 'ole'-:type do. z=. z, type biff_supbook source
elseif. 'dde'-:type do. z=. z, type biff_supbook source
end.
for_ni. extname do.
  z=. z, type biff_externname >ni
end.
for_ni. crn do.
  z=. z, type biff_crn >ni
end.
z
)

coclass 'biffxf'
coinsert 'biff'
getcolor=: ]
getfont=: 3 : 0
'fontheight fontitalic fontstrike fontcolor fontweight fontscript fontunderline fontfamily fontcharset fontname'=. y
y=. fontheight ; fontitalic ; fontstrike ; (getcolor fontcolor) ; fontweight ; fontscript ; fontunderline ; fontfamily ; fontcharset ; dltb fontname
if. (#fontset__COCREATOR)= n=. fontset__COCREATOR i. y do.
  fontset__COCREATOR=: fontset__COCREATOR, y
end.
n
)

getformat=: 3 : 0
if. (#formatset__COCREATOR)= n=. formatset__COCREATOR i. <y=. dltb y do.
  formatset__COCREATOR=: formatset__COCREATOR, <y
end.
n
)
getxfrow=: 3 : 0
font=. getfont fontheight ; fontitalic ; fontstrike ; fontcolor ; fontweight ; fontscript ; fontunderline ; fontfamily ; fontcharset ; fontname
formatn=. getformat format
if. (0=leftlinecolor) *. leftlinestyle do. leftlinecolor=. 16b40 end.
if. (0=rightlinecolor) *. rightlinestyle do. rightlinecolor=. 16b40 end.
if. (0=toplinecolor) *. toplinestyle do. toplinecolor=. 16b40 end.
if. (0=bottomlinecolor) *. bottomlinestyle do. bottomlinecolor=. 16b40 end.
if. (0=diagonalcolor) *. diagonalstyle do. diagonalcolor=. 16b40 end.
if. (0=patternbgcolor) *. pattern do. patternbgcolor=. 16b41 end.
typeprotparent=. lock bitor 1 bitshl hideformula bitor 1 bitshl type bitor 2 bitshl parentxf
align=. horzalign bitor 3 bitshl textwrap bitor 1 bitshl vertalign
indentshrink=. indent bitor 4 bitshl shrink
used=. 2 bitshl usedformat bitor 1 bitshl usedfont bitor 1 bitshl usedalign bitor 1 bitshl usedborder bitor 1 bitshl usedbackground bitor 1 bitshl usedprotect
border=. leftlinestyle bitor 4 bitshl rightlinestyle bitor 4 bitshl toplinestyle bitor 4 bitshl bottomlinestyle bitor 4 bitshl (getcolor leftlinecolor) bitor 7 bitshl (getcolor rightlinecolor) bitor 7 bitshl diagonaltopleft bitor 1 bitshl diagonalbottomleft
linecolor=. (getcolor toplinecolor) bitor 7 bitshl (getcolor bottomlinecolor) bitor 7 bitshl (getcolor diagonalcolor) bitor 7 bitshl diagonalstyle bitor 5 bitshl pattern
color=. (getcolor patterncolor) bitor 7 bitshl (getcolor patternbgcolor)
font, formatn, typeprotparent, align, rotation, indentshrink, used, border, linecolor, color
)
copyxfrow=: 3 : 0
'font formatn typeprotparent align rotate indentshrink used border linecolor color'=. y
'fontheight fontitalic fontstrike fontcolor fontweight fontscript fontunderline fontfamily fontcharset fontname'=: font{fontset__COCREATOR
format=: >formatn{formatset__COCREATOR
lock=: 1 bitand typeprotparent
hideformula=: _1 bitshl 2b10 bitand typeprotparent
type=: _2 bitshl 2b100 bitand typeprotparent
parentxf=: _4 bitshl 16bfff0 bitand typeprotparent
rotation=: rotate
horzalign=: 7 bitand align
textwrap=: _3 bitshl 8 bitand align
vertalign=: _4 bitshl 16b70 bitand align
indent=: 16bf bitand indentshrink
shrink=: _4 bitshl 16b10 bitand indentshrink
'usedformat usedfont usedalign usedborder usedbackground usedprotect'=: |. _2}. _8{. (8#0), #: used
leftlinestyle=: 16bf bitand border
rightlinestyle=: _4 bitshl 16bf0 bitand border
toplinestyle=: _8 bitshl 16bf00 bitand border
bottomlinestyle=: _12 bitshl 16bf000 bitand border
leftlinecolor=: _16 bitshl 16b7f0000 bitand border
rightlinecolor=: _23 bitshl 16b3f800000 bitand border
diagonaltopleft=: _30 bitshl 16b40000000 bitand border
diagonalbottomleft=: _31 bitshl (DFH '0x80000000') bitand border
toplinecolor=: 16b7f bitand linecolor
bottomlinecolor=: _7 bitshl 16b3f80 bitand linecolor
diagonalcolor=: _14 bitshl 16b1fc000 bitand linecolor
diagonalstyle=: _21 bitshl 16b1e00000 bitand linecolor
pattern=: _26 bitshl (DFH '0xfc000000') bitand linecolor
patterncolor=: 16b7f bitand color
patternbgcolor=: _7 bitshl 16b3f80 bitand color
)
copyxfobj=: 3 : 0
l=. y
nm=. nl__l 0
for_nmi. nm do. (>nmi)=: ((>nmi), '__l')~ end.
)

create=: 3 : 0
'fontheight fontitalic fontstrike fontcolor fontweight fontscript fontunderline fontfamily fontcharset fontname'=: {.fontset_COCREATOR
format=: 'General'
lock=: 0
hideformula=: 1
type=: 0  
parentxf=: 0
horzalign=: 0  
textwrap=: 0
vertalign=: 2  
rotation=: 0
indent=: 0
shrink=: 0
usedformat=: 0
usedfont=: 0
usedalign=: 0
usedborder=: 0
usedbackground=: 0
usedprotect=: 0
leftlinestyle=: 0
rightlinestyle=: 0
toplinestyle=: 0
bottomlinestyle=: 0
leftlinecolor=: 0
rightlinecolor=: 0
diagonaltopleft=: 0
diagonalbottomleft=: 0
toplinecolor=: 0
bottomlinecolor=: 0
diagonalcolor=: 0
diagonalstyle=: 0
pattern=: 0
patterncolor=: colorborder
patternbgcolor=: colorpattern
xfindex=: _1   
if. ''-.@-:y do.
  if. (3!:0 y) e. 1 4 do.
    copyxfrow y
  elseif. (3!:0 y) e. 32 do.
    copyxfobj y
  end.
end.
)

destroy=: codestroy

coclass 'biffsheet'
coinsert 'biff'
insertpicture=: 4 : 0
img=. y
'rowcol xyoffset scalexy'=. x
if. 2 131072 e.~ 3!:0 rowcol do. rowcol=. A1toRC rowcol end.
'row col'=. rowcol
'x1 y1'=. xyoffset
'scalex scaley'=. scalexy
z=. processbitmap img
if. _1&= 0{::z do. z return. end.
'width height size data'=. }.z
width=. width * scalex
height=. height * scaley
positionImage row, col, x1, y1, width, height
record=. 16b007f
if. RECORDLEN>:8 + size do.
  length=. 8 + size
  cf=. 16b09
  env=. 16b01
  lcb=. size
  header=. toWORD0 record, length, cf, env
  header=. header, toDWORD0 lcb
  imdata=: imdata, header, data
else.
  length=. RECORDLEN
  cf=. 16b09
  env=. 16b01
  lcb=. size
  header=. toWORD0 record, length, cf, env
  header=. header, toDWORD0 lcb
  imdata=: imdata, header, (RECORDLEN-8){.data
  data=. (RECORDLEN-8)}.data
  a=. <.(#data) % RECORDLEN
  b=. RECORDLEN | #data
  if. #a do.
    imdata=: imdata,, (toHeader 16b003c, RECORDLEN) ,("1) (a, RECORDLEN)$(a*RECORDLEN){.data
  end.
  if. #b do.
    imdata=: imdata, (toHeader 16b003c, b), (-b){.data
  end.
end.
0 ; ''
)
position_object=: 3 : 0
'rowstart colstart x1 y1 width height'=. y
colend=. colstart  
rowend=. rowstart  
if. x1 >: sizeCol colstart do. x1=. 0 end.
if. y1 >: sizeRow rowstart do. y1=. 0 end.
width=. width + x1
height=. height + y1
while. width >: sizeCol colend do.
  width=. width - sizeCol colend
  colend=. >:colend
end.
while. height >: sizeRow rowend do.
  height=. height - sizeRow rowend
  rowend=. >:rowend
end.
if. 0= sizeCol colstart do. return. end.
if. 0= sizeCol colend do. return. end.
if. 0= sizeRow rowstart do. return. end.
if. 0= sizeRow rowend do. return. end.
x1=. <. 1024 * x1 % sizeCol colstart
y1=. <. 256 * y1 % sizeRow rowstart
x2=. <. 1024 * width % sizeCol colend 
y2=. <. 256 * height % sizeRow rowend 
colstart , x1 , rowstart , y1 , colend , x2 , rowend , y2
)

positionImage=: storeobjpicture@:position_object

getcolsizes=: 3 : 0
if. (#colsizes)= i=. y i.~ {.("1) colsizes do.
  _1
else.
  {: i{colsizes
end.
)

getrowsizes=: 3 : 0
if. (#rowsizes)= i=. y i.~ {.("1) rowsizes do.
  _1
else.
  {: i{rowsizes
end.
)
sizeCol=: 3 : 0
if. _1~: getcolsizes y do.
  if. 0= getcolsizes y do.
    0 return.
  else.
    <. 5 + 8 * 0.38 + 256%~ getcolsizes y return.
  end.
else.
  <. 5 + 8 * 0.38 + 256%~ defaultcolwidth * 256 return.
end.
)
sizeRow=: 3 : 0
if. _1~: getrowsizes y do.
  if. 0= getrowsizes y do.
    0 return.
  else.
    <. (4%3) * 20%~ getrowsizes y return.
  end.
else.
  <. (4%3) * 20%~ defaultrowheight return.
end.
)
storeobjpicture=: 3 : 0
'colL dxL rwT dyT colR dxR rwB dyB'=. y
record=. 16b005d   
length=. 16b003c   
cObj=. 16b0001   
OT=. 16b0008   
id=. 16b0001   
grbit=. 16b0614   
cbMacro=. 16b0000   
Reserved1=. 16b0000   
Reserved2=. 16b0000   
icvBack=. 16b09     
icvFore=. 16b09     
fls=. 16b00     
fAuto=. 16b00     
icv=. 16b08     
lns=. 16bff     
lnw=. 16b01     
fAutoB=. 16b00     
frs=. 16b0000   
cf=. 16b0009   
Reserved3=. 16b0000   
cbPictFmla=. 16b0000   
Reserved4=. 16b0000   
grbit2=. 16b0001   
Reserved5=. 16b0000   
header=. toWORD0 record, length
data=. toDWORD0 cObj
data=. data, toWORD0 OT
data=. data, toWORD0 id
data=. data, toWORD0 grbit
data=. data, toWORD0 colL
data=. data, toWORD0 dxL
data=. data, toWORD0 rwT
data=. data, toWORD0 dyT
data=. data, toWORD0 colR
data=. data, toWORD0 dxR
data=. data, toWORD0 rwB
data=. data, toWORD0 dyB
data=. data, toWORD0 cbMacro
data=. data, toDWORD0 Reserved1
data=. data, toWORD0 Reserved2
data=. data, toBYTE icvBack
data=. data, toBYTE icvFore
data=. data, toBYTE fls
data=. data, toBYTE fAuto
data=. data, toBYTE icv
data=. data, toBYTE lns
data=. data, toBYTE lnw
data=. data, toBYTE fAutoB
data=. data, toWORD0 frs
data=. data, toDWORD0 cf
data=. data, toWORD0 Reserved3
data=. data, toWORD0 cbPictFmla
data=. data, toWORD0 Reserved4
data=. data, toWORD0 grbit2
data=. data, toDWORD0 Reserved5
imdata=: imdata, header, data
)
processbitmap=: 3 : 0
raiseError=. ''
if. 32=3!:0 y do.
  data=. ]`(''"_)@.(_1&-:)@:fread y
else.
  data=. y
end.
if. ((#data) <: 16b36) do.
  raiseError=. 'size error'
  goto_error.
end.
identity=. 2{.data
if. (identity -.@-: 'BM') do.
  raiseError=. 'signiture error'
  goto_error.
end.
data=. 2}.data
size=. fromDWORD0 4{.data
data=. 4}.data
size=. size - 16b36 
size=. size + 16b0c 
data=. 12}.data
'width height'=. fromDWORD0 8{.data
data=. 8}.data
if. (width > 16bffff) do.
  raiseError=. 'bitmap width greater than 65535'
  goto_error.
end.
if. (height > 16bffff) do.
  raiseError=. 'bitmap height greater than 65535'
  goto_error.
end.
planesandbitcount=. fromWORD0 4{.data
data=. 4}.data
if. (24 ~: 1{planesandbitcount) do.  
  raiseError=. 'not 24 bit color'
  goto_error.
end.
if. (1 ~: 0{planesandbitcount) do.
  raiseError=. 'contain more than 1 plane'
  goto_error.
end.
compression=. fromDWORD0 4{.data
data=. 4}.data
if. (compression ~: 0) do.
  raiseError=. 'compression not supported'
  goto_error.
end.
data=. 20}.data
header=. toDWORD0 16b000c
header=. header, toWORD0 width, height, 16b01, 16b18
data=. header, data
0 ; width ; height ; size ; data
return.
label_error.
_1 ; raiseError
)

insertbackground=: 3 : 0
z=. processbitmap y
if. _1&= 0{::z do. z return. end.
'width height size data'=. }.z
record=. 16b00e9
if. RECORDLEN>:8 + size do.
  length=. 8 + size
  cf=. 16b09
  env=. 16b01
  lcb=. size
  header=. toWORD0 record, length, cf, env
  header=. header, toDWORD0 lcb
  imdata=: imdata, header, data
else.
  length=. RECORDLEN
  cf=. 16b09
  env=. 16b01
  lcb=. size
  header=. toWORD0 record, length, cf, env
  header=. header, toDWORD0 lcb
  imdata=: imdata, header, (RECORDLEN-8){.data
  data=. (RECORDLEN-8)}.data
  a=. <.(#data) % RECORDLEN
  b=. RECORDLEN | #data
  if. #a do.
    imdata=: imdata,, (toHeader 16b003c, RECORDLEN) ,("1) (a, RECORDLEN)$(a*RECORDLEN){.data
  end.
  if. #b do.
    imdata=: imdata, (toHeader 16b003c, b), (-b){.data
  end.
end.
0 ; ''
)

adjdim=: 3 : 0
dimensions=: ((1 3{dimensions)>. y) 1 3}dimensions
)

initsheet=: 3 : 0
sheetname=: y
calccount=: 100
calcmode=: 1
refmode=: 1
delta=: 0.001000
iteration=: 0
printheaders=: 0
printgridlines=: 0
defaultrowheightnotmatch=: 1
defaultrowheighthidden=: 0
defaultrowheightspaceabove=: 0
defaultrowheightspacebelow=: 0
defaultrowheight=: 315  
wsbool=: 0
horizontalpagebreaks=: 0 3$''         
verticalpagebreaks=: 0 3$''           
header=: ''  
footer=: 'Page &P of &N'
hcenter=: 0
vcenter=: 0
leftmargin=: 0.5
rightmargin=: 0.5
topmargin=: 0.5
bottommargin=: 0.75
'setuppapersize setupscaling setupstartpage setupfitwidth setupfitheight setuprowmajor setupportrait setupinvalid setupblackwhite setupdraft setupcellnote setuporientinvalid setupusestartpage setupnoteatend setupprinterror setupdpi setupvdpi setupheadermargin setupfootermargin setupnumcopy'=: 0
setuprowmajor=: 1
setupportrait=: 1
setupinvalid=: 0
setupscaling=: 100
setupdpi=: setupvdpi=: 600
setupheadermargin=: 0.75
setupfootermargin=: 0.75
backgroundbitmap=: ''
protect=: 1
windowprotect=: 1
objectprotect=: 1
scenprotect=: 1
password=: ''
defaultcolwidth=: 8
colinfoset=: 0 7$''
dimensions=: 0 0 0 0
window2=: 16b6b6 0 0 10
sclnum=: sclden=: 1
'xsplit ysplit topvis leftvis'=: 0
activepane=: ''
mergedcell=: 0 4$''               
rowlabelrange=: collabelrange=: 0 4$''
condformatstream=: ''
selection=: 0 5$''
hlink=: 0 8$''
imdata=: ''
dvstream=: ''
colsizes=: 0 2$''
rowsizes=: 0 2$''
sheetoffset=: 0
object_ids=: 1024 0 0 1024
images_array=: 0 0$''
image_mso_size=: 0
comments_array=: 0 0$''
charts_array=: 0 0$''
comments_author=: ''
comments_visible=: 0
filter_area=: ''
filter_count=: 0
filter_on=: 0
filter_cols=: 0 0$''
rowcolused=: 0 2$0
)

writestream=: 3 : 0
z=. biff_bof 16b600, 16
p1=. #z
z=. z, biff_index (0 1+2{.dimensions), 0
z=. z, biff_calccount calccount
z=. z, biff_calcmode calcmode
z=. z, biff_refmode refmode
z=. z, biff_delta delta
z=. z, biff_iteration iteration
z=. z, biff_printheaders printheaders
z=. z, biff_printgridlines printgridlines
z=. z, biff_defaultrowheight defaultrowheightnotmatch, defaultrowheighthidden, defaultrowheightspaceabove, defaultrowheightspacebelow, defaultrowheight
z=. z, biff_wsbool wsbool
z=. z, biff_horizontalpagebreaks horizontalpagebreaks
z=. z, biff_verticalpagebreaks verticalpagebreaks
z=. z, biff_header header
z=. z, biff_footer footer
z=. z, biff_hcenter hcenter
z=. z, biff_vcenter vcenter
z=. z, biff_leftmargin leftmargin
z=. z, biff_rightmargin rightmargin
z=. z, biff_topmargin topmargin
z=. z, biff_bottommargin bottommargin
z=. z, biff_setup setuppapersize ; setupscaling ; setupstartpage ; setupfitwidth ; setupfitheight ; setuprowmajor ; setupportrait ; setupinvalid ; setupblackwhite ; setupdraft ; setupcellnote ; setuporientinvalid ; setupusestartpage ; setupnoteatend ; setupprinterror ; setupdpi ; setupvdpi ; setupheadermargin ; setupfootermargin ; setupnumcopy
z=. z, backgroundbitmap
if. (#password) *. protect +. windowprotect +. objectprotect +. scenprotect do.
  if. protect do. z=. z, biff_protect protect end.
  if. windowprotect do. z=. z, biff_windowprotect windowprotect end.
  if. objectprotect do. z=. z, biff_objectprotect objectprotect end.
  if. scenprotect do. z=. z, biff_scenprotect scenprotect end.
  z=. z, biff_password passwordhash password
end.
p2=. #z
z=. (toDWORD0 p2+y) (p1+16+i.4)}z
z=. z, biff_defaultcolwidth defaultcolwidth
for_item. colinfoset do. z=. z, ({.item) biff_colinfo }.item end.

z=. z, store_filtermode ''
z=. z, store_autofilterinfo ''
z=. z, store_autofilters ''

z=. z, biff_dimensions dimensions
z=. z, stream, imdata
z=. z, store_images ''
z=. z, store_charts ''
z=. z, store_filters ''
z=. z, store_comments ''

z=. z, biff_window2 window2
z=. z, biff_scl sclnum, sclden
if. #activepane do. z=. z, biff_pane (xsplit, ysplit) ; (topvis, leftvis) ; activepane end.
if. #selection do. z=. z, biff_selection selection end.
if. (2:~:$$) mergedcell do. mergedcell=: _4]\, mergedcell end.
z=. z, biff_mergedcells mergedcell
z=. z, biff_labelranges rowlabelrange ; collabelrange
z=. z, condformatstream
for_item. hlink do. z=. z, biff_hlink item end.
z=. z, dvstream
z=. z, biff_eof ''
)
passwordhash=: 3 : 0
pw=. (15<.#y){.y
hash=. 0 [ i=. 0
while. i<#pw do.
  c=. 3&u: u: i{pw
  i=. >:i
  hash=. hash bitxor #. i&|. _15{.#: c
end.
hash=. 16bce4b bitxor (#pw) bitxor hash
)

create=: 3 : 0
stream=: ''
initsheet y
)

destroy=: codestroy

coclass 'biffbook'
coinsert 'biff'
getxfobj=: 3 : 0
z=. _1  
if. ''-:y do.
  z=. cxf
else.
  if. (3!:0 y) e. 32 do.
    z=. y
  elseif. (3!:0 y) e. 1 4 do.
    if. 1=#y do.
      z=. addxfobj ({.y){xfset
    elseif. (}.$xfset)-:$y do.
      if. (#xfset)= n=. xfset i. y do.
        xfset=: xfset, y
      end.
      z=. addxfobj n{xfset
    end.
  end.
end.
z
)

getxfidx=: 3 : 0
rc=. 0
if. ''-:y do. y=. cxf end.
if. (3!:0 y) e. 32 do.
  l=. y
  if. (#xfset)= n=. xfset i. a=. getxfrow__l '' do.
    xfset=: xfset, a
    xfindex__l=: n
  end.
  rc=. n
elseif. (3!:0 y) e. 1 4 do.
  if. 1=#y do.
    rc=. {.y
  elseif. (}.$xfset)-:$y do.
    if. (#xfset)= n=. xfset i. y do.
      xfset=: xfset, y
    end.
    rc=. n
  end.
end.
rc
)

add2sst=: 3 : 0
sstn=: >:sstn
if. (#sst) = b=. sst i. y do. sst=: sst, y end.
b
)

WriteSST=: 3 : 0
sstbuf=: (toWORD0 16b00fc, 0), toDWORD0 sstn, #sst
wrtp=: 0 [ bufn=: RECORDLEN-8
for_ix. sst do.
  oix=. >ix
  if. 131072= 3!:0 oix do.
    if. bufn<5 do. wrtcont '' end.
    wrtn #oix
    wrtw oix
  else.
    if. bufn<4 do. wrtcont '' end.
    wrtn #oix
    wrt oix
  end.
end.
sstbuf=: (toWORD0 RECORDLEN-bufn) (wrtp+2+i.2)}sstbuf
)

wrtn=: 3 : 0
sstbuf=: sstbuf, toWORD0 y
bufn=: bufn-2
)

wrtw=: 3 : 0
while. #y do.
  if. bufn<1+2*#y do.
    sstbuf=: sstbuf, (1{a.), toucode0 (<.-:bufn-1){.y
    y=. (<.-:bufn-1)}.y
    bufn=: bufn - 1+ 2*(<.-:bufn-1)
    wrtcont ''
  else.
    sstbuf=: sstbuf, (1{a.), toucode0 y
    bufn=: bufn - 1+2*#y
    y=. ''
  end.
end.
)

wrt=: 3 : 0
while. #y do.
  if. bufn<1+#y do.
    sstbuf=: sstbuf, (0{a.), (bufn-1){.y
    y=. (bufn-1)}.y
    bufn=: bufn - 1+ bufn-1
    wrtcont ''
  else.
    sstbuf=: sstbuf, (0{a.), y
    bufn=: bufn - 1+#y
    y=. ''
  end.
end.
)

wrtcont=: 3 : 0
sstbuf=: (toWORD0 RECORDLEN-bufn) (wrtp+2+i.2)}sstbuf
wrtp=: #sstbuf
sstbuf=: sstbuf, toWORD0 16b003c, 0
bufn=: RECORDLEN
)

initbook=: 3 : 0
fileprotectionstream=: ''
workbookprotectionstream=: ''
codepage=: 1200
window1=: 672 192 11004 6636 16b38
backup=: 0
hideobj=: 0
date1904=: 0
precision=: 1
bookbool=: 1
fontset=: 0 0$''
fontset=: fontset, fontsize ; 0 ; 0 ; colorfont ; 400 ; 0 ; 0 ; 0 ; 0 ; fontname     
fontset=: fontset, fontsize ; 0 ; 0 ; colorfont ; 400 ; 0 ; 0 ; 0 ; 0 ; fontname     
fontset=: fontset, fontsize ; 0 ; 0 ; colorfont ; 400 ; 0 ; 0 ; 0 ; 0 ; fontname     
fontset=: fontset, fontsize ; 0 ; 0 ; colorfont ; 400 ; 0 ; 0 ; 0 ; 0 ; fontname     
fontset=: fontset, fontsize ; 0 ; 0 ; colorfont ; 400 ; 0 ; 0 ; 0 ; 0 ; fontname     
fontset=: fontset, fontsize ; 0 ; 0 ; colorfont ; 400 ; 0 ; 0 ; 0 ; 0 ; fontname     
formatset=: format0n$<''
formatset=: ('General';'0';'0.00';'#,##0';'#,##0.00';'"$"#,##0_); ("$"#,##0)';'"$"#,##0_);[Red] ("$"#,##0)';'"$"#,##0.00_); ("$"#,##0.00)';'"$"#,##0.00_);[Red] ("$"#,##0.00)';'0%';'0.00%';'0.00E+00';'# ?/?';'# ??/??';'M/D/YY';'D-MMM-YY';'D-MMM';'MMM-YY';'h:mm AM/PM';'h:mm:ss AM/PM';'h:mm';'h:mm:ss';'M/D/YY h:mm') (i.23) } formatset
formatset=: ('_(#,##0_);(#,##0)';'_(#,##0_);[Red](#,##0)';'_(#,##0.00_);(#,##0.00)';'_(#,##0.00_);[Red](#,##0.00)';'_ ("$"* #,##0_);_ ("$"* (#,##0);_ ("$"* "-"_);_(@_)';'_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)';'_ ("$"* #,##0.00_);_ ("$"* (#,##0.00);_ ("$"* "-"??_);_(@_)';'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';'mm:ss';'[h]:mm:ss';'mm:ss.0';'##0.0E+0';'@') (37+i.13)}formatset
formatset=: formatset, 'd/m/yyyy';'#,##0.000';'#,##0.0000';'#,##0.000000000'
xfset=: 0 10$''
xfset=: xfset, 0 0 16bfff5 16b20 0 0 0 0 0 16b20c0 
xfset=: xfset, 1 0 16bfff5 16b20 0 0 16bf4 0 0 16b20c0 
xfset=: xfset, 1 0 16bfff5 16b20 0 0 16bf4 0 0 16b20c0
xfset=: xfset, 2 0 16bfff5 16b20 0 0 16bf4 0 0 16b20c0
xfset=: xfset, 2 0 16bfff5 16b20 0 0 16bf4 0 0 16b20c0
xfset=: xfset, 0 0 16bfff5 16b20 0 0 16bf4 0 0 16b20c0
xfset=: xfset, 0 0 16bfff5 16b20 0 0 16bf4 0 0 16b20c0
xfset=: xfset, 0 0 16bfff5 16b20 0 0 16bf4 0 0 16b20c0
xfset=: xfset, 0 0 16bfff5 16b20 0 0 16bf4 0 0 16b20c0
xfset=: xfset, 0 0 16bfff5 16b20 0 0 16bf4 0 0 16b20c0
xfset=: xfset, 0 0 16bfff5 16b20 0 0 16bf4 0 0 16b20c0
xfset=: xfset, 0 0 16bfff5 16b20 0 0 16bf4 0 0 16b20c0
xfset=: xfset, 0 0 16bfff5 16b20 0 0 16bf4 0 0 16b20c0
xfset=: xfset, 0 0 16bfff5 16b20 0 0 16bf4 0 0 16b20c0
xfset=: xfset, 0 0 16bfff5 16b20 0 0 16bf4 0 0 16b20c0
xfset=: xfset, 0 0 1 16b20 0 0 0 0 0 16b20c0 
xfset=: xfset, 1 16b2b 16bfff5 16b20 0 0 16bf8 0 0 16b20c0 
xfset=: xfset, 1 16b29 16bfff5 16b20 0 0 16bf8 0 0 16b20c0
xfset=: xfset, 1 16b2c 16bfff5 16b20 0 0 16bf8 0 0 16b20c0
xfset=: xfset, 1 16b2a 16bfff5 16b20 0 0 16bf8 0 0 16b20c0
xfset=: xfset, 1 9 16bfff5 16b20 0 0 16bf8 0 0 16b20c0
biffxfset=: ''
styleset=: 0 5$''
styleset=: styleset, 16b10 ; 1 ; 3 ; 16bff ; ''
styleset=: styleset, 16b11 ; 1 ; 6 ; 16bff ; ''
styleset=: styleset, 16b12 ; 1 ; 4 ; 16bff ; ''
styleset=: styleset, 16b13 ; 1 ; 7 ; 16bff ; ''
styleset=: styleset, 16b00 ; 1 ; 0 ; 16bff ; ''
styleset=: styleset, 16b14 ; 1 ; 5 ; 16bff ; ''
colorset=: 0 16bffffff 16bff 16bff00 16bff0000 16bffff 16bff00ff 16bffff00
colorset=: colorset, RGB 16b00 16b00 16b00  
colorset=: colorset, RGB 16bff 16bff 16bff  
colorset=: colorset, RGB 16bff 16b00 16b00  
colorset=: colorset, RGB 16b00 16bff 16b00  
colorset=: colorset, RGB 16b00 16b00 16bff  
colorset=: colorset, RGB 16bff 16bff 16b00  
colorset=: colorset, RGB 16bff 16b00 16bff  
colorset=: colorset, RGB 16b00 16bff 16bff  
colorset=: colorset, RGB 16b80 16b00 16b00  
colorset=: colorset, RGB 16b00 16b80 16b00  
colorset=: colorset, RGB 16b00 16b00 16b80  
colorset=: colorset, RGB 16b80 16b80 16b00  
colorset=: colorset, RGB 16b80 16b00 16b80  
colorset=: colorset, RGB 16b00 16b80 16b80  
colorset=: colorset, RGB 16bc0 16bc0 16bc0  
colorset=: colorset, RGB 16b80 16b80 16b80  
colorset=: colorset, RGB 16b99 16b99 16bff  
colorset=: colorset, RGB 16b99 16b33 16b66  
colorset=: colorset, RGB 16bff 16bff 16bcc  
colorset=: colorset, RGB 16bcc 16bff 16bff  
colorset=: colorset, RGB 16b66 16b00 16b66  
colorset=: colorset, RGB 16bff 16b80 16b80  
colorset=: colorset, RGB 16b00 16b66 16bcc  
colorset=: colorset, RGB 16bcc 16bcc 16bff  
colorset=: colorset, RGB 16b00 16b00 16b80  
colorset=: colorset, RGB 16bff 16b00 16bff  
colorset=: colorset, RGB 16bff 16bff 16b00  
colorset=: colorset, RGB 16b00 16bff 16bff  
colorset=: colorset, RGB 16b80 16b00 16b80  
colorset=: colorset, RGB 16b80 16b00 16b00  
colorset=: colorset, RGB 16b00 16b80 16b80  
colorset=: colorset, RGB 16b00 16b00 16bff  
colorset=: colorset, RGB 16b00 16bcc 16bff  
colorset=: colorset, RGB 16bcc 16bff 16bff  
colorset=: colorset, RGB 16bcc 16bff 16bcc  
colorset=: colorset, RGB 16bff 16bff 16b99  
colorset=: colorset, RGB 16b99 16bcc 16bff  
colorset=: colorset, RGB 16bff 16b99 16bcc  
colorset=: colorset, RGB 16bcc 16b99 16bff  
colorset=: colorset, RGB 16bff 16bcc 16b99  
colorset=: colorset, RGB 16b33 16b66 16bff  
colorset=: colorset, RGB 16b33 16bcc 16bcc  
colorset=: colorset, RGB 16b99 16bcc 16b00  
colorset=: colorset, RGB 16bff 16bcc 16b00  
colorset=: colorset, RGB 16bff 16b99 16b00  
colorset=: colorset, RGB 16bff 16b66 16b00  
colorset=: colorset, RGB 16b66 16b66 16b99  
colorset=: colorset, RGB 16b96 16b96 16b96  
colorset=: colorset, RGB 16b00 16b33 16b66  
colorset=: colorset, RGB 16b33 16b99 16b66  
colorset=: colorset, RGB 16b00 16b33 16b00  
colorset=: colorset, RGB 16b33 16b33 16b00  
colorset=: colorset, RGB 16b99 16b33 16b00  
colorset=: colorset, RGB 16b99 16b33 16b66  
colorset=: colorset, RGB 16b33 16b33 16b99  
colorset=: colorset, RGB 16b33 16b33 16b33  
country=: 1 1
supbook=: ''
externsheet=: 0 3$''
extname=: ''
refname=: ''
datasize=: 0
biffsize=: 0
mso_size=: 0
mso_clusters=: 0 5$''
images_size=: 0
images_data=: 0 0$''  
)

addhlink=: 3 : 0
l=. sheeti{sheet
hlink__l=: hlink__l, y
)

writefileprotection=: 3 : 0
fileprotectionstream=: y
)

writeworkbookprotection=: 3 : 0
workbookprotectionstream=: y
)

writecelltable=: 3 : 0
l=. sheeti{sheet
stream__l=: stream__l, y
)

writecondformattable=: 3 : 0
l=. sheeti{sheet
condformatstream__l=: y
)

writedvtable=: 3 : 0
l=. sheeti{sheet
dvstream__l=: y
)

celladr=: 4 : 0
'relrow relcol'=. x
'row col'=. y
z=. toWORD0 row
z=. z, toWORD0 (16bff bitand col) bitor 14 bitshl (0~:relcol) bitor 1 bitshl 0~:relrow
)

cellrangeadr=: 4 : 0
'relrow1 relrow2 relcol1 relcol2'=. x
'row1 row2 col1 col2'=. y
z=. toWORD0 row1, row2
z=. z, toWORD0 (16bff bitand col1) bitor 14 bitshl (0~:relcol1) bitor 1 bitshl 0~:relrow1
z=. z, toWORD0 (16bff bitand col2) bitor 14 bitshl (0~:relcol2) bitor 1 bitshl 0~:relrow2
)

newsupbook=: 3 : 0
supbook=: supbook, y conew 'biffsupbook'
supbooki=: <:#supbook
)

newrefname=: 3 : 0
refname=: refname, '' conew 'biffrefname'
{:refname
)

newextname=: 3 : 0
l=. supbooki{supbook
newextname__l y
)

newcrn=: 3 : 0
l=. supbooki{supbook
newcrn__l y
)

create=: 3 : 0
if. ''-:y do.
  'fontname fontsize'=: ((IFUNIX+IFWINE){::'Courier New' ; 'Monospace') ; 220
  sheetname=. 'Sheet1'
else.
  'fontname fontsize'=: 2{.y
  sheetname=. 0{:: (2}.y) , <'Sheet1'
end.
sstn=: #sst=: ''
xfset=: sheet=: ''
initbook ''
addsheet sheetname  
cxf=: addxfobj 15{xfset  
)

destroy=: 3 : 0
if. DEBUG do.
  for_l. sheet do.
    if. n=. #rowcolused__l do.
      if. n ~: # a=. ~.rowcolused__l do.
        1!:2&2 sheetname__l, ' duplicated cell'
        1!:2&2 (1 < rowcolused__l #/. rowcolused__l) # a
      end.
    end.
  end.
end. for_l. sheet do. destroy__l '' end.
for_l. biffxfset do. destroy__l '' end.
for_l. supbook do. destroy__l '' end.
for_l. refname do. destroy__l '' end.
codestroy ''
)
rgbcolor=: 3 : 0
(i: <./) +/("1) | y -("1) RGBtuple colorset
)
addxfobj=: 3 : 0
biffxfset=: biffxfset, z=. y conew 'biffxf'
z
)
addsheet=: 3 : 0
sheet=: sheet, ((y-:''){:: y ; 'Sheet', ": >:#sheet) conew 'biffsheet'
sheeti=: <:#sheet
)
save=: 3 : 0
fn=. >y
calc_mso_sizes ''
z=. biff_bof 16b600, 5
z=. z, fileprotectionstream
z=. z, biff_codepage codepage
z=. z, workbookprotectionstream
z=. z, biff_window1 window1
z=. z, biff_backup backup
z=. z, biff_hideobj hideobj
z=. z, biff_date1904 date1904
z=. z, biff_precision precision
z=. z, biff_bookbool bookbool
fontset1=. (4{.fontset), 5}.fontset
for_item. fontset1 do. z=. z, biff_font item end.
if. 164<#formatset do.
  for_i. i.164-~#formatset do. z=. z, biff_format (i+164) ; (i+164){formatset end.
end.
for_item. xfset do. z=. z, biff_xf item end.
for_item. styleset do. z=. z, (0{::item) biff_style }.item end.
if. 8<#colorset do.
  z=. z, biff_palette 8}.colorset
end.
olesheet=. ''
olehead=. z
seekpoint=. #z
z=. ''
for_item. sheet do.
  z=. z, a=. biff_boundsheet 0 ; 0 ; 0 ; sheetname__item
  seekpoint=. seekpoint, #a
end.
z=. z, biff_country country
z=. z, add_mso_drawing_group ''
z=. z, WriteSST ''
if. #supbook do.
  for_item. supbook do.
    z=. z, writestream__item #sheet
  end.
  z=. z, biff_externsheet externsheet
  for_item. refname do.
    z=. z, writestream__item ''
  end.
end.
z=. z, biff_eof ''
olehead=. olehead, z
sheetoffset=. ({.seekpoint)+ #z
for_item. sheet do.
  z=. writestream__item {:sheetoffset
  olesheet=. olesheet, z
  sheetoffset=. sheetoffset, #z
end.
seekpoint=. }:+/\seekpoint
sheetoffset=. }:+/\sheetoffset
for_i. i.#seekpoint do.
  p1=. 4+i{seekpoint
  p2=. 0
  z=. toDWORD0 i{sheetoffset
  olehead=. z (p1+i.4)}olehead
end.
stream=. ('Workbook' ; '' ; '') conew 'oleppsfile'
append__stream olehead
append__stream olesheet
root=. (0 ; 0 ; <stream) conew 'oleppsroot'
rc=. save__root fn ; 0 ; ''
destroy__root ''
destroy__stream ''
rc
)
addcolinfo=: 3 : 0
cxf addcolinfo y
:
'col1 col2 width hide level collapse'=. 6{.y
l=. sheeti{sheet
colsizes__l=: ~. colsizes__l, (col1 + i.>:col2-col1) ,("0) (0~:hide){width, 0
colinfoset__l=: colinfoset__l, (getxfidx x), col1, col2, width, hide, level, collapse
''
)
addrowinfo=: 3 : 0
cxf addrowinfo y
:
'rownumber firstcol lastcols usedefaultheight rowheight'=. 5{. y=. 12{.y
l=. sheeti{sheet
if. 0=usedefaultheight do.
  rowsizes__l=: ~. rowsizes__l, rownumber, rowheight
end.
stream__l=: stream__l, (getxfidx x) biff_row y
''
)
writestring=: 3 : 0
cxf writestring y
:
assert. 2 3 e.~ #y
assert. 1 2 4 8 131072 e.~ (3!:0) 0{::y
assert. 2 32 131072 e.~ (3!:0) 1{::y
if. 2 131072 e.~ 3!:0 rc=. 0{::y do. y=. (<A1toRC rc) 0}y end.
l=. sheeti{sheet
xf=. getxfidx x
if. 3=#y do. opt=. 2{::y else. opt=. 0 end.
if. (0=opt) *. 0 e. $yn=. 1{::y do. '' return. end.  
if. 2 131072 e.~ 3!:0 yn do.
  if. 2> $$yn do.
    adjdim__l 0{::y
    stream__l=: stream__l, xf biff_label 2{.y
    if. DEBUG do.
      rowcolused__l=: rowcolused__l, ,0{::y
    end.
    '' return.
  elseif. 2=$$yn do.
    yn=. ,. <@dtb("1) yn    
  elseif. 3=$$yn do.
    yn=. <@dtb("1) yn       
  elseif. do. 'unhandled exception' 13!:8 (3)
  end.
elseif. -. 32 e.~ 3!:0 yn do.
  'unhandled exception' 13!:8 (3)
end.
if. 1=#@, yn do.  
  if. (0=opt) *. (0:=#), >yn do. '' return. end. 
  adjdim__l 0{::y
  stream__l=: stream__l, xf biff_label ({.y), <, >yn
  if. DEBUG do.
    rowcolused__l=: rowcolused__l, ,0{::y
  end.
elseif. 3>$$yn do.
  if. 1=$$yn do. yn=. ,:yn end.
  s=. $yn
  'r c'=. 0{::y
  if. (0=opt) *. 0= +./ f1=. ,f=. (<'') ~: yn do. '' return. end.
  if. 0=opt do.
    mr=. 1 i.~ 1 e.("1) f
    mc=. 1 i.~ 1 e.("1) |:f
    adjdim__l (mr, mc) + 0{::y
    mr=. 1 i:~ 1 e.("1) f
    mc=. 1 i:~ 1 e.("1) |:f
    adjdim__l (mr, mc) + 0{::y
  else.
    adjdim__l 0{::y
    adjdim__l (<:s) + 0{::y
  end.
  sst=: sst, (~.f1#,yn) -. sst
  sstn=: sstn + +/f1
  stream__l=: stream__l,, (toHeader 16b00fd, 10) ,("1) (_2]\ toWORD0 f1#({:s)#r+i.{.s) ,("1) (_2]\ toWORD0 f1#,({.s)#,:c+i.{:s) ,("1) (toWORD0 xf) ,("1) (_4]\ toDWORD0 sst i. f1#,yn)
  if. DEBUG do.
    rowcolused__l=: rowcolused__l, (f1#({:s)#r+i.{.s) ,. f1#,({.s)#,:c+i.{:s
  end.
  if. (1=opt) *. 0 e. f1 do.
    stream__l=: stream__l,, (toHeader 16b0201, 6) ,("1) (_2]\ toWORD0 (-.f1)#({:s)#r+i.{.s) ,("1) (_2]\ toWORD0 (-.f1)#,({.s)#,:c+i.{:s) ,("1) (toWORD0 xf)
  end.
elseif. do. 'unhandled exception' 13!:8 (3)
end.
''
)
writeblank=: 3 : 0
cxf writeblank y
:
assert. 1 2 4 8 131072 e.~ (3!:0) y
if. 2 131072 e.~ 3!:0 y do. y=. A1toRC y end.
assert. 0=2|#,y
if. 0=#,y do. '' return. end.  
'r c'=. |: _2]\ ,y
l=. sheeti{sheet
xf=. getxfidx x
adjdim__l (<./r), <./c
adjdim__l (>./r), >./c
stream__l=: stream__l,, (toHeader 16b0201, 6) ,("1) (_2]\ toWORD0 r) ,("1) (_2]\ toWORD0 c) ,("1) (toWORD0 xf)
if. DEBUG do.
  rowcolused__l=: rowcolused__l, r ,. c
end.
''
)
writeinteger=: 3 : 0
cxf writeinteger y
:
assert. 2 3 e.~ #y
assert. 1 2 4 8 131072 e.~ (3!:0) 0{::y
assert. 1 4 8 e.~ (3!:0) 1{::y
if. 2 131072 e.~ 3!:0 rc=. 0{::y do. y=. (<A1toRC rc) 0}y end.
l=. sheeti{sheet
xf=. getxfidx x
if. 536870911 < >./ |, 1{::y do. x writenumber y end.
if. (0:=#), yn=. 2b10 bitor 2 bitshl <. 1{::y do. '' return. end.  
if. 3=#y do. opt=. 2{::y else. opt=. 0 end.
if. 1=#@, yn do.  
  if. (1=opt) *. 0= <. 1{::y do. '' return. end.
  adjdim__l 0{::y
  stream__l=: stream__l, xf biff_integer ({.y), < {., 1{::y
  if. DEBUG do.
    rowcolused__l=: rowcolused__l, ,0{::y
  end.
elseif. 3>$$yn do.
  if. 1=$$yn do. yn=. ,:yn end.
  s=. $yn
  'r c'=. 0{::y
  if. 0~:opt do.
    if. (1=opt) *. 0= +./ f1=. ,f=. 0~: <. 1{::y do. '' return. end.
    if. 1=opt do.
      mr=. 1 i.~ 1 e.("1) f
      mc=. 1 i.~ 1 e.("1) |:f
      adjdim__l (mr, mc) + 0{::y
      mr=. 1 i:~ 1 e.("1) f
      mc=. 1 i:~ 1 e.("1) |:f
      adjdim__l (mr, mc) + 0{::y
    else.
      adjdim__l 0{::y
      adjdim__l (<:s) + 0{::y
    end.
    stream__l=: stream__l,, (toHeader 16b027e, 10) ,("1) (_2]\ toWORD0 f1#({:s)#r+i.{.s) ,("1) (_2]\ toWORD0 f1#,({.s)#,:c+i.{:s) ,("1) (toWORD0 xf) ,("1) (_4]\ toDWORD0 f1#,yn)
    if. DEBUG do.
      rowcolused__l=: rowcolused__l, (f1#({:s)#r+i.{.s) ,. f1#,({.s)#,:c+i.{:s
    end.
    if. 2=opt do.
      stream__l=: stream__l,, (toHeader 16b0201, 6) ,("1) (_2]\ toWORD0 (-.f1)#({:s)#r+i.{.s) ,("1) (_2]\ toWORD0 (-.f1)#,({.s)#,:c+i.{:s) ,("1) (toWORD0 xf)
    end.
  else.
    adjdim__l 0{::y
    adjdim__l (<:s) + 0{::y
    stream__l=: stream__l,, (toHeader 16b027e, 10) ,("1) (_2]\ toWORD0 ({:s)#r+i.{.s) ,("1) (_2]\ toWORD0 ,({.s)#,:c+i.{:s) ,("1) (toWORD0 xf) ,("1) (_4]\ toDWORD0 ,yn)
    if. DEBUG do.
      rowcolused__l=: rowcolused__l, (({:s)#r+i.{.s) ,. ,({.s)#,:c+i.{:s
    end.
  end.
elseif. do. 'unhandled exception' 13!:8 (3)
end.
''
)
writenumber=: 3 : 0
cxf writenumber y
:
assert. 2 3 e.~ #y
assert. 1 2 4 8 131072 e.~ (3!:0) 0{::y
assert. 1 4 8 e.~ (3!:0) 1{::y
if. 2 131072 e.~ 3!:0 rc=. 0{::y do. y=. (<A1toRC rc) 0}y end.
l=. sheeti{sheet
xf=. getxfidx x
if. (0:=#), yn=. 1{::y do. '' return. end.  
if. 3=#y do. opt=. 2{::y else. opt=. 0 end.
if. 1=#@, yn do.  
  if. (1=opt) *. 0=1{::y do. '' return. end.
  adjdim__l 0{::y
  stream__l=: stream__l, xf biff_number ({.y), < {., yn
  if. DEBUG do.
    rowcolused__l=: rowcolused__l, ,0{::y
  end.
elseif. 3>$$yn do.
  if. 1=$$yn do. yn=. ,:yn end.
  s=. $yn
  'r c'=. 0{::y
  if. 0~:opt do.
    if. (1=opt) *. 0= +./ f1=. ,f=. 0~: <. 1{::y do. '' return. end.
    if. 1=opt do.
      mr=. 1 i.~ 1 e.("1) f
      mc=. 1 i.~ 1 e.("1) |:f
      adjdim__l (mr, mc) + 0{::y
      mr=. 1 i:~ 1 e.("1) f
      mc=. 1 i:~ 1 e.("1) |:f
      adjdim__l (mr, mc) + 0{::y
    else.
      adjdim__l 0{::y
      adjdim__l (<:s) + 0{::y
    end.
    stream__l=: stream__l,, (toHeader 16b0203, 14) ,("1) (_2]\ toWORD0 f1#({:s)#r+i.{.s) ,("1) (_2]\ toWORD0 f1#,({.s)#,:c+i.{:s) ,("1) (toWORD0 xf) ,("1) (_8]\ toDouble0 f1#,yn)
    if. DEBUG do.
      rowcolused__l=: rowcolused__l, (f1#({:s)#r+i.{.s) ,. f1#,({.s)#,:c+i.{:s
    end.
    if. 2=opt do.
      stream__l=: stream__l,, (toHeader 16b0201, 6) ,("1) (_2]\ toWORD0 (-.f1)#({:s)#r+i.{.s) ,("1) (_2]\ toWORD0 (-.f1)#,({.s)#,:c+i.{:s) ,("1) (toWORD0 xf)
    end.
  else.
    adjdim__l 0{::y
    adjdim__l s + 0{::y
    stream__l=: stream__l,, (toHeader 16b0203, 14) ,("1) (_2]\ toWORD0 ({:s)#r+i.{.s) ,("1) (_2]\ toWORD0 ,({.s)#,:c+i.{:s) ,("1) (toWORD0 xf) ,("1) (_8]\ toDouble0 ,yn)
    if. DEBUG do.
      rowcolused__l=: rowcolused__l, (({:s)#r+i.{.s) ,. ,({.s)#,:c+i.{:s
    end.
  end.
elseif. do. 'unhandled exception' 13!:8 (3)
end.
''
)
writenumber2=: 3 : 0
cxf writenumber2 y
:
assert. 3 4 e.~ #y
assert. 1 2 4 8 131072 e.~ (3!:0) 0{::y
assert. 1 4 8 e.~ (3!:0) 1{::y
assert. 2 131072 e.~ (3!:0) 2{::y
if. 2 131072 e.~ 3!:0 rc=. 0{::y do. y=. (<A1toRC rc) 0}y end.
l=. getxfobj x
t=. format__l
format__l=: 2{::y
if. 8= (3!:0) 1{::y do.
  l writenumber (2{.y), (4=#y)#{:y
else.
  l writeinteger (2{.y), (4=#y)#{:y
end.
format__l=: t
''
)
writedate=: 3 : 0
cxf writedate y
:
assert. 2 3 e.~ #y
assert. 1 2 4 8 131072 e.~ (3!:0) 0{::y
assert. 1 4 8 e.~ (3!:0) 1{::y
if. 2 131072 e.~ 3!:0 rc=. 0{::y do. y=. (<A1toRC rc) 0}y end.
l=. getxfobj x
t=. format__l
if. 2=#y do. y=. y, <shortdatefmt end. 
assert. 2 131072 e.~ (3!:0) 2{::y
format__l=: 2{::y
l writenumber ({.y), <36522-~ 1{::y
format__l=: t
''
)
insertpicture=: 4 : 0
l=. sheeti{sheet
x insertpicture__l y
)
insertbackground=: 3 : 0
l=. sheeti{sheet
insertbackground__l y
)
printarea=: 4 : 0
if. _1=x do.
  x=. newsupbook 'internal' ; #sheet
end.
externsheet=: externsheet, x, 0 0
r=. newrefname ''
builtin__r=: 1
name__r=: 16b06{a.
sheet__r=: 0
formula__r=: (a.{~16b3b 0 0), (0 0 0 0 cellrangeadr y)
x
)
rowrepeat=: 4 : 0
if. _1=x do.
  x=. newsupbook 'internal' ; #sheet
end.
externsheet=: externsheet, x, 0 0
r=. newrefname ''
builtin__r=: 1
name__r=: 16b07{a.
sheet__r=: {.y
formula__r=: (a.{~16b3b 0 0), 0 0 0 0 cellrangeadr, (}.y), 0 255
x
)
colrepeat=: 4 : 0
if. _1=x do.
  x=. newsupbook 'internal' ; #sheet
end.
externsheet=: externsheet, x, 0 0
r=. newrefname ''
builtin__r=: 1
name__r=: 16b07{a.
sheet__r=: {.y
formula__r=: (a.{~16b3b 0 0), 0 0 0 0 cellrangeadr 0 16bffff, (}.y)
x
)
rowcolrepeat=: 4 : 0
if. _1=x do.
  x=. newsupbook 'internal' ; #sheet
end.
externsheet=: externsheet, x, 0 0
r=. newrefname ''
builtin__r=: 1
name__r=: 16b07{a.
sheet__r=: {.y
formula__r=: (16b29{a.), (toWORD0@# , ]) ((a.{~16b3b 0 0), 0 0 0 0 cellrangeadr 0 16bffff, (>_2{.}.y)), ((a.{~16b3b 0 0), 0 0 0 0 cellrangeadr, (>2{.}.y), 0 255), (16b10{a.)
x
)
coclass 'biffread'
coinsert 'oleutlfcn'
create=: 4 : 0
debug=: x
stream=: y
biffver=: 0        
streamtype=: 16    
filepos=: 0
worksheets=: 0 2$''
bkrecords=: 0 3$'' [ bkbytes=: ''   
records=: 0 3$'' [ bytes=: ''       
sst=: ''
bnsst=: nsst=: 0  
sstchar=: 0
sstrtffe=: 0
insst=. 0
type=. _1
newptr=. 0
while. (type~:10)*.filepos<#stream do.
  'type ptr len'=. nextrecord ''
  if. debug do.
    bkbytes=: bkbytes, (ptr+i.len){stream
    newptr=. newptr+len [ bkrecords=: bkrecords, type, newptr, len
  end.
  if. 16b9=type do.        
    biffver=: 16b200 break.
  elseif. 16b209=type do.  
    biffver=: 16b300 break.
  elseif. 16b409=type do.  
    biffver=: 16b400 break.
  elseif. 16b809=type do.  
    if. 0=biffver do.  
      'biffver streamtype'=: fromWORD0 4{.data=. (ptr+i.len){stream
    end.
  elseif. 16b0085=type do.  
    if. 0={.fromBYTE 5{data=. (ptr+i.len){stream do.
      if. 16b500=biffver do.
        worksheets=: worksheets, (0 decodestring8 6}.data) ,< {.(fromDWORD0 4{.data)
      else.
        worksheets=: worksheets, (0 decodeustring8 6}.data) ,< {.(fromDWORD0 4{.data)
      end.
    end.
  elseif. 16b00fc=type do.  
    'bnsst nsst'=: fromDWORD0 8{.data=. (ptr+i.len){stream
    readsst 8}.data
    insst=. 1
  elseif. insst *. 16b003c=type do. 
    readsst data=. (ptr+i.len){stream
  elseif. do.
    insst=. 0
  end.
end.
if. #sst do. sst=: <;._2 sst end.
assert. nsst=#sst
assert. 0~:biffver
)

destroy=: codestroy
readsst=: 3 : 0
fp=. 0
while. fp<#y do.
  if. 0=sstchar+sstrtffe do.
    't t1'=. fp decodeustring16 y
    sst=: sst, t, {.a.
    fp=. fp + a [ 'a b c'=. t1
    sstchar=: b-#t
    sstrtffe=: c
  else.
    't t1'=. (fp, sstchar, sstrtffe) decodeustring16a y
    sst=: (}:sst), t, {.a.
    fp=. fp + a [ 'a b c'=. t1
    sstchar=: sstchar-#t
    sstrtffe=: c
  end.
end.
)
nextrecord=: 3 : 0
'type len'=. fromWORD0 (filepos+i.4){stream
filepos=: +/data=. (4+filepos), len
type, data
)
decodestring8=: 4 : 0
len=. {.fromBYTE (x+0){y
< len (([ <. #@]) {. ]) (x+1)}.y
)

decodestring16=: 4 : 0
len=. {.fromWORD0 (x+i.2){y
< len (([ <. #@]) {. ]) (x+2)}.y
)
decodeustring8=: 4 : 0
len=. {.fromBYTE (x+0){y
uc=. 0~:1 bitand op=. {.fromBYTE (x+1){y
fe=. 0~:4 bitand op
rtf=. 0~:8 bitand op
if. uc do.
  < fromucode0 (2*len) (([ <. #@]) {. ]) (x+2+(fe*4)+(rtf*2))}.y
else.
  < len (([ <. #@]) {. ]) (x+2+(fe*4)+(rtf*2))}.y
end.
)

decodeustring16=: 4 : 0
len=. {.fromWORD0 (x+i.2){y
uc=. 0~:1 bitand op=. {.fromBYTE (x+2){y  
fe=. 0~:4 bitand op                         
rtf=. 0~:8 bitand op                        
lenrtffe=. 0
if. rtf do. lenrtffe=. 4* {.fromWORD0 ((x+3)+i.2){y end.
if. fe do. lenrtffe=. lenrtffe + {.fromDWORD0 ((x+3+(rtf*2))+i.4){y end.
l=. (3+(fe*4)+(rtf*2)) + (len*uc{1 2) + lenrtffe  
if. uc do.
  z=. fromucode0 z1=. (2*len) (([ <. #@]) {. ]) (x+3+(fe*4)+(rtf*2))}.y
  z2=. (2*len)-#z1  
else.
  z=. z1=. len (([ <. #@]) {. ]) (x+3+(fe*4)+(rtf*2))}.y
  z2=. len-#z1  
end.
if. (#y)<x+l do.
  z ; (#y), len, (x+l)-z2+#y
else.
  z ; l, len, 0
end.
)
decodeustring16a=: 4 : 0
'x len lenrtffe'=. x
if. len do.
  uc=. 0~:1 bitand op=. {.fromBYTE (x+0){y
  l=. 1 + (len*uc{1 2) + lenrtffe  
  if. uc do.
    z=. fromucode0 z1=. (2*len) (([ <. #@]) {. ]) (x+1)}.y
    z2=. (2*len)-#z1    
  else.
    z=. z1=. len (([ <. #@]) {. ]) (x+1)}.y
    z2=. len-#z1  
  end.
else.
  l=. lenrtffe  
  z=. ''
  z2=. 0
end.
if. (#y)<x+l do.
  z ; (#y), len, (x+l)-z2+#y
else.
  z ; l, len, 0
end.
)
readsheet=: 0&$: : (4 : 0)
if. 16b200 16b300 16b400 e.~ biffver do.  
  filepos=: 0
  scnt=. _1
  sheetfound=. 0
else.
  filepos=: y
  sheetfound=. 1
end.
rowcolss=. rowcol=. rowcol4=. rowcol8=. rowcolc=. 0 2$''
cellvalss=. cellval=. cellval4=. cellval8=. cellvalc=. ''
null=. {.a.
type=. _1
records=: 0 3$'' [ bytes=: ''       
newptr=. 0
lookstr=. 0
while. filepos<#stream do.
  'type ptr len'=. nextrecord ''
  if. debug do.
    bytes=: bytes, (ptr+i.len){stream
    newptr=. newptr+len [ records=: records, type, newptr, len
  end.
  if. 0=sheetfound do.  
    if. 16b0009 16b0209 16b0409 e.~ type do.  
      if. 16b10= fromWORD0 (2+ptr+i.2){stream do.  
        if. y = scnt=. >:scnt do.         
          sheetfound=. 1
        end.
      end.
    end.
    continue.
  end.
  select. type
  case. 16b000a do. 
    break.
  case. 16b027e do. 
    if. 0=x do.
      if. 8=3!:0 a=. >getrk 4}.data=. (ptr+i.len){stream do.
        rowcol8=. rowcol8, fromWORD0 4{.data
        cellval8=. cellval8, a
      else.
        rowcol4=. rowcol4, fromWORD0 4{.data
        cellval4=. cellval4, a
      end.
    else.
      rowcol=. rowcol, fromWORD0 4{.data=. (ptr+i.len){stream
      cellval=. cellval, ; null&,@ (":(!.maxpp))&.> getrk 4}.data   
    end.
  case. 16b0002 do. 
    if. 0=x do.
      rowcol4=. rowcol4, fromWORD0 4{.data=. (ptr+i.len){stream
      cellval4=. cellval4, {.fromDWORD0 (2#{.a.),~ (7+i.2){data
    else.
      rowcol=. rowcol, fromWORD0 4{.data=. (ptr+i.len){stream
      cellval=. cellval, null&,@": {.fromDWORD0 (2#{.a.),~ (7+i.2){data
    end.
  case. 16b0003 do. 
    if. 0=x do.
      rowcol8=. rowcol8, fromWORD0 4{.data=. (ptr+i.len){stream
      cellval8=. cellval8, {.fromDouble0 (7+i.8){data
    else.
      rowcol=. rowcol, fromWORD0 4{.data=. (ptr+i.len){stream
      cellval=. cellval, null&,@ (":(!.maxpp)) {.fromDouble0 (7+i.8){data
    end.
  case. 16b0203 do. 
    if. 0=x do.
      rowcol8=. rowcol8, fromWORD0 4{.data=. (ptr+i.len){stream
      cellval8=. cellval8, {.fromDouble0 (6+i.8){data
    else.
      rowcol=. rowcol, fromWORD0 4{.data=. (ptr+i.len){stream
      cellval=. cellval, null&,@ (":(!.maxpp)) {.fromDouble0 (6+i.8){data
    end.
  case. 16b00fd do. 
    rowcolss=. rowcolss, fromWORD0 4{.data=. (ptr+i.len){stream
    cellvalss=. cellvalss, fromDWORD0 (6+i.4){data
  case. 16b0005 do. 
    if. ({.a.)=(8+ptr){stream do. 
      if. 0=x do.
        rowcol4=. rowcol4, fromWORD0 4{.data=. (ptr+i.len){stream
        cellval4=. cellval4, ({.a.)~:7{data
      else.
        rowcol=. rowcol, fromWORD0 4{.data=. (ptr+i.len){stream
        cellval=. cellval, null&,@": ({.a.)~:7{data
      end.
    end.
  case. 16b0205 do. 
    if. ({.a.)=(7+ptr){stream do. 
      if. 0=x do.
        rowcol4=. rowcol4, fromWORD0 4{.data=. (ptr+i.len){stream
        cellval4=. cellval4, ({.a.)~:6{data
      else.
        rowcol=. rowcol, fromWORD0 4{.data=. (ptr+i.len){stream
        cellval=. cellval, null&,@": ({.a.)~:6{data
      end.
    end.
  case. 16b00bd do. 
    rocol=. fromWORD0 4{.data=. (ptr+i.len){stream
    lc=. fromWORD0 _2{.data
    nc=. >:lc-{:rocol
    if. 0=x do.
      v=. getrk("1) _6]\ _2}.4}.data
      for_rcl. rocol +("1) 0 ,("0) i.nc do.
        if. 8=3!:0 a=. >rcl_index{v do.
          rowcol8=. rowcol8, rcl
          cellval8=. cellval8, a
        else.
          rowcol4=. rowcol4, rcl
          cellval4=. cellval4, a
        end.
      end.
    else.
      rowcol=. rowcol, rocol +("1) 0 ,("0) i.nc
      cellval=. cellval, ; null&,@ (":(!.maxpp))&.> getrk("1) _6]\ _2}.4}.data
    end.
  case. 16b0006 do. 
    data=. (ptr+i.len){stream
    if. 16b200 = biffver do. 
      if. (255 255{a.)-:13 14{data do.
        if. (0{a.)=7{data do. 
          lookrowcol=. fromWORD0 4{.data
          lookstr=. 2
        elseif. (1{a.)=7{data do. 
          if. 0=x do.
            rowcol4=. rowcol4, fromWORD0 4{.data
            cellval4=. cellval4, ({.a.)~:9{data
          else.
            rowcol=. rowcol, fromWORD0 4{.data
            cellval=. cellval, null&,@": ({.a.)~:9{data
          end.
        end.
      else.  
        if. 0=x do.
          rowcol8=. rowcol8, fromWORD0 4{.data
          cellval8=. cellval8, {.fromDouble0 (7+i.8){data
        else.
          rowcol=. rowcol, fromWORD0 4{.data
          cellval=. cellval, null&,@ (":(!.maxpp)) {.fromDouble0 (7+i.8){data
        end.
      end.
    else.
      if. (255 255{a.)-:12 13{data do.
        if. (0{a.)=6{data do. 
          lookrowcol=. fromWORD0 4{.data
          lookstr=. 2
        elseif. (1{a.)=6{data do. 
          if. 0=x do.
            rowcol4=. rowcol4, fromWORD0 4{.data
            cellval4=. cellval4, ({.a.)~:8{data
          else.
            rowcol=. rowcol, fromWORD0 4{.data
            cellval=. cellval, null&,@": ({.a.)~:8{data
          end.
        end.
      else.  
        if. 0=x do.
          rowcol8=. rowcol8, fromWORD0 4{.data
          cellval8=. cellval8, {.fromDouble0 (6+i.8){data
        else.
          rowcol=. rowcol, fromWORD0 4{.data
          cellval=. cellval, null&,@ (":(!.maxpp)) {.fromDouble0 (6+i.8){data
        end.
      end.
    end.
  case. 16b0007 do. 
    if. 1=lookstr do.
      if. 0=x do.
        rowcolc=. rowcolc, lookrowcol
        cellvalc=. cellvalc, null&,@> {.0 decodestring8 data=. (ptr+i.len){stream
      else.
        rowcol=. rowcol, lookrowcol
        cellval=. cellval, null&,@> {.0 decodestring8 data=. (ptr+i.len){stream
      end.
    end.
  case. 16b0207 do.
    if. 1=lookstr do.
      if. 16b300 16b400 16b500 e.~ biffver do.  
        if. 0=x do.
          rowcolc=. rowcolc, lookrowcol
          cellvalc=. cellvalc, null&,@> {.0 decodestring16 data=. (ptr+i.len){stream
        else.
          rowcol=. rowcol, lookrowcol
          cellval=. cellval, null&,@> {.0 decodestring16 data=. (ptr+i.len){stream
        end.
      else.
        if. 0=x do.
          rowcolc=. rowcolc, lookrowcol
          cellvalc=. cellvalc, null&,@> {.0 decodeustring16 data=. (ptr+i.len){stream
        else.
          rowcol=. rowcol, lookrowcol
          cellval=. cellval, null&,@> {.0 decodeustring16 data=. (ptr+i.len){stream
        end.
      end.
    end.
  case. 16b0004 do. 
    if. 0=x do.
      rowcolc=. rowcolc, fromWORD0 4{.data=. (ptr+i.len){stream
      cellvalc=. cellvalc, null&,@> {.0 decodestring8 7}.data
    else.
      rowcol=. rowcol, fromWORD0 4{.data=. (ptr+i.len){stream
      cellval=. cellval, null&,@> {.0 decodestring8 7}.data
    end.
  case. 16b0204 do. 
    if. 16b300 16b400 16b500 e.~ biffver do.  
      if. 0=x do.
        rowcolc=. rowcolc, fromWORD0 4{.data=. (ptr+i.len){stream
        cellvalc=. cellvalc, null&,@> {.0 decodestring16 6}.data
      else.
        rowcol=. rowcol, fromWORD0 4{.data=. (ptr+i.len){stream
        cellval=. cellval, null&,@> {.0 decodestring16 6}.data
      end.
    else.
      if. 0=x do.
        rowcolc=. rowcolc, fromWORD0 4{.data=. (ptr+i.len){stream
        cellvalc=. cellvalc, null&,@> {.0 decodeustring16 6}.data
      else.
        rowcol=. rowcol, fromWORD0 4{.data=. (ptr+i.len){stream
        cellval=. cellval, null&,@> {.0 decodeustring16 6}.data
      end.
    end.
  case. 16b00d6 do. 
    if. 16b300 16b400 16b500 e.~ biffver do.  
      if. 0=x do.
        rowcolc=. rowcolc, fromWORD0 4{.data=. (ptr+i.len){stream
        cellvalc=. cellvalc, null&,@> {.0 decodestring16 6}.data
      else.
        rowcol=. rowcol, fromWORD0 4{.data=. (ptr+i.len){stream
        cellval=. cellval, null&,@> {.0 decodestring16 6}.data
      end.
    else.
      if. 0=x do.
        rowcolc=. rowcolc, fromWORD0 4{.data=. (ptr+i.len){stream
        cellvalc=. cellvalc, null&,@> {.0 decodeustring16 6}.data
      else.
        rowcol=. rowcol, fromWORD0 4{.data=. (ptr+i.len){stream
        cellval=. cellval, null&,@> {.0 decodeustring16 6}.data
      end.
    end.
  end.
  lookstr=. 0>.<:lookstr
end.
if. 0=x do.
  (rowcol4, rowcol8, rowcolc, rowcolss) (;<) (<("0) cellval4), (<("0) cellval8), ( <;._1 cellvalc), sst{~cellvalss
else.
  (rowcol, rowcolss) (;<) ( <;._1 cellval), sst{~cellvalss
end.
)
getrk=: 3 : 0
if. 0=2 bitand d=. fromDWORD0 2}.y do. 
  bigendian=: ({.a.)={. 1&(3!:4) 1  
  if. 0=bigendian do.
    rk=. fromDouble0 toDWORD0 0, d bitand DFH '0xfffffffc'
  else.
    rk=. fromDouble0 toDWORD0 0,~ d bitand DFH '0xfffffffc'
  end.
else.  
  rk=. _2 bitsha d bitand DFH '0xfffffffc'
end.
if. 1 bitand d do.  
  rk=. rk%100
end.
<{.rk
)
readxlsheets=: 3 : 0
0 readxlsheets y
:
try.
  'fln strng'=. 2{.!.(<0) boxopen y
  x=. boxopen x
  locs=. '' 
  (msg=. 'file not found') assert fexist fln
  locs=. locs,ole=. fln conew 'olestorage'
  if. 0=#wks=. getppssearch__ole 'Workbook' ; 1 ; 1 do.              
    if. 0=#wks=. getppssearch__ole 'Book' ; 1 ; 1 do.                
      (msg=. 'unknown Excel file format') assert 16b40009 16b60209 16b60409 e.~ fromDWORD0 freadx fln;0 4  
    end.
  end.
  locs=. locs,wks
  locs=. locs,ex=. conew 'biffread'
  if. #wks do.
    wk=. {.wks
    0&create__ex data__wk
    if. x-:<'' do.
      x=. i.#worksheets__ex
    elseif. -. */ 1 4 8 e.~ 3!:0 every x do.
      x=. x i.&(tolower&.>"_)~ {."1 worksheets__ex   
    elseif. do. x=. >x 
    end.
    (msg=. 'worksheet not found') assert x<#worksheets__ex
    'name location'=. |: x{worksheets__ex
  else.
    0&create__ex fread fln
    location=. {.>x
  end.
  'ix cell'=. <"1 |: ,:^:(1:=#@:$) strng&readsheet__ex"0 location 
  for_l. |.locs do. 
    destroy__l ''
    locs=. locs -. l
  end.
  rcs=. 0&>.@(>./ >:@- <./) each ix   
  offset_biffread_=: >{: off=. <./ each ix   
  ix=. ix -"1 each off
  m=. 0$0
  for_sht. cell do.
    n=. (sht_index{::rcs)$a:
    n=. (>sht) (<"1^:(0:<#) sht_index{::ix)} n
    m=. m,<n
  end.
  dtb=. #~ ([: +./\. ' '&~:)
  (<@dtb"1 name) ,. m
catch.
  for_l. |.locs do. 
    destroy__l ''
    locs=. locs -. l
  end.
  smoutput 'readxlsheets: ',msg
end.
)
stringtype=:  (<1) ,~ [: {. boxopen  
firstsheet=: 0: :({.^:(3!:0 -.@e. 2 131072"_)@[) 
readxlsheetsstring=: 0&$: : ([ readxlsheets stringtype@])
readxlworkbook=: ''&readxlsheets
readexcel=: 0 _1 {:: firstsheet readxlsheets ]
readexcelstring=: 0 _1 {:: firstsheet readxlsheets stringtype@]
readxlsheetnames=: 3 : 0
try.
  fln=. boxopen y
  locs=. '' 
  (msg=. 'file not found') assert fexist fln
  locs=. locs,ole=. fln conew 'olestorage'
  if. 0=#wks=. getppssearch__ole 'Workbook' ; 1 ; 1 do.              
    if. 0=#wks=. getppssearch__ole 'Book' ; 1 ; 1 do.                
      (msg=. 'unknown Excel file format') assert 16b40009 16b60209 16b60409 e.~ fromDWORD0 freadx fln;0 4  
    end.
  end.
  locs=. locs,wks
  locs=. locs,ex=. conew 'biffread'
  if. #wks do.
    wk=. {.wks
    0&create__ex data__wk
  else.
    0&create__ex fread y
  end.
  nms=. {."1 worksheets__ex
  for_l. |.locs do. destroy__l '' end. 
  nms
catch.
  for_l. |.locs do. destroy__l '' end.
  smoutput 'readxlsheetnames: ',msg
end.
)
dumpexcel=: 0&$: : (4 : 0)
assert. fexist y
ole=. (>y) conew 'olestorage'
if. 0=#wks=. getppssearch__ole 'Workbook' ; 1 ; 1 do.              
  if. 0=#wks=. getppssearch__ole 'Book' ; 1 ; 1 do.                
    assert. 16b40009 16b60209 16b60409 e.~ fromDWORD0 freadx y;0 4  
  end.
end.
ex=. conew 'biffread'
if. #wks do.
  wk=. {.wks
  1&create__ex data__wk       
  if. 2 131072 e.~ 3!:0 x do. x=. (<x) i.~ {.("1) worksheets__ex end.
  assert. x<#worksheets__ex
  'name location'=. x{worksheets__ex
else.
  1&create__ex fread y
  location=. x
end.
'ix cell'=. 0&readsheet__ex location
sst__=: sst__ex
bkrecords__=: bkrecords__ex
bkbytes__=: bkbytes__ex
records__=: records__ex
bytes__=: bytes__ex
destroy__ex ''
for_wk. wks do. destroy__wk '' end.
destroy__ole ''
''
)
readexcel_z_=: readexcel_biffread_
readexcelstring_z_=: readexcelstring_biffread_
dumpexcel_z_=: dumpexcel_biffread_
readxl_z_=: readexcel_biffread_
readxlstring_z_=: readexcelstring_biffread_
readxlsheetnames_z_=: readxlsheetnames_biffread_
readxlsheets_z_=: readxlsheets_biffread_
readxlsheetsstring_z_=: readxlsheetsstring_biffread_
readxlworkbook_z_=: readxlworkbook_biffread_
coclass 'biffwrite'
coinsert 'biffbook'
writexlsheets=: 4 : 0
  try.
    locs=. '' 
    if. 0=#x do. empty'' return. end. 
    (msg=. 'too many levels of boxing') assert 2>:L. x
    msg=. 'error in left argument'
    shts=. makexarg x
    shtnme=. ((0 < #) {:: 'Sheet1'&;) (<0 0) {:: shts
    locs=. locs,bi=. ('Arial';220;shtnme) conew 'biffbook'
    shtdat=. (<0 1){:: shts
    msg=. 'error writing data for ',shtnme
    bi writeshtdat shtdat 
    shts=. }.shts         
    msg=. 'error creating/writing later sheets'
    if. #shts do. bi addsheets"1 shts end. 
    binary=. save__bi y
    success=. destroy__bi ''
    (*#binary){:: success;binary
  catch.
    for_l. |.locs do. 
      destroy__l ''
      locs=. locs -. l
    end.
    0 [ smoutput 'writexlsheets: ',msg
  end.
)
mfv1=: ,:^:(#&$ = 1:)       
mfva=: ,:^:([: 2&> #@$)^:_  
ischar=: 3!:0 e. 2 131072"_
firstones=: > 0: , }:
lastones=: > 0: ,~ }.
indices=: 4:$.$.       
isrowblks=: >/@(+/^:(#&$)"_1) 
isdata=: 3 : 0
  if. 1< lvls=. L. y do. 0 return. 
  elseif. 0= lvls do. 1 return. 
  elseif. 2< cols=. {:$ y do. 1 return. 
  elseif. (-. *./(ischar&> +. a:&=) {."1 y) *. 2 = cols do. 1 return. 
  elseif. 1< #@$ &> {:"1 y do. 0 return. 
  elseif. do. 1 
  end.
)
makexarg=: 3 : 0
  if. isdata y do.
    if. 2= 3!:0 y do. y=. <y end.
    y=. a:,. <mfva y
  else.
    if. 2>{:$ y do. 
      y=. a:,.y
    end.
    if. #idx=. (I. b=. ischar &>{:"1 y) do.  
      upd=. ({."1 ,. <&.>@({:"1)) b#y 
      y=. upd idx }y
    end.
    if. #idx=. (I. b=. 2> #@$&>{:"1 y) do.  
      upd=. ({."1 ,. mfva&.>@({:"1)) b#y 
      y=. upd idx }y
    end.
  end.
  mfv1 y
)
addsheets=: 4 : 0
  'shtnme shtdat'=. y
  addsheet__x shtnme
  x writeshtdat shtdat
)
writeshtdat=: 4 : 0
  if. 0=L.y do.
    writenumber__x 0 0;y
  else.
    as=. ischar &> y
    blks=. blocks as
    tls=. {.0 2|: blks
    dat=. blks <;.0 y 
    writestring__x"1 (<"1 tls),.dat
    blks=. blocks -.as
    tls=. {.0 2|: blks
    dat=. blks ([:<>);.0 y  
    writenumber__x"1 (<"1 tls),.dat
  end.
)
tlc=: [: I. firstones      
brc=: [: I. lastones       
tlbrc=: <@(tlc ,. brc)"1   
bpr=: i.@# ,:"0 1&.> tlbrc 

mtch=: 4 : 0
  's t'=. <"0 x (](#~; (#~-.)) e.~&:(<@{:"2))&> {.y
  t=. t((,&.>{:)`[)@.(1=#@])y
  s=. x([:(<@{:"2 ({:@{.@{:(<0 1)} {.)/.]) ,)&> s
  s;t
)
blcks=: [:|:"2 [:; [:mtch/ bpr

tlshape=: ([,: (-~>:))/"_1 
blocks=: tlshape@:blcks  
writexlsheets_z_=: writexlsheets_biffwrite_
coclass 'biffmd4'

md4=: crc32_md4=: 3 : 0
z=. 2&(3!:4) 4{. 128!:3 y
assert. 16=#z
z
)

coclass 'biffsheet'
coinsert 'biff'
embedchart=: 4 : 0
chart=. y
'rowcol xy_offset xy_scale'=. x
if. 2 131072 e.~ 3!:0 rowcol do. rowcol=. A1toRC rowcol end.
charts_array=: charts_array, rowcol ; chart ; xy_offset ; xy_scale
)
insertimage=: 4 : 0
img=. y
'rowcol xy_offset xy_scale'=. x
if. 2 131072 e.~ 3!:0 rowcol do. rowcol=. A1toRC rowcol end.
if. 32=3!:0 img do.
  if. -.fexist img do. 'file not exist' 13!:8 (3) end.
  images_array=: images_array, rowcol ; (>img) ; xy_offset ; xy_scale ; (16#{.a.) ; 0 ; 0 ; 0 0
else.
  img=. (1&u: ::]) img  
  assert. 2=3!:0 img
  images_array=: images_array, rowcol ; img ; xy_offset ; xy_scale ; (16#{:a.) ; 0 ; 0 ; 0 0
end.
''
)
writecomment=: 4 : 0

rowcol=. >@{. y=. boxopen y
opt=. }.y
string=. x
if. 2 131072 e.~ 3!:0 rowcol do. rowcol=. A1toRC rowcol end.
if. 2|#,opt do. 'Uneven number of additional arguments' 13!:8 (3) end.
comments_array=: comments_array, comment_params (rowcol ; string), opt
''
)
prepare_images=: 3 : 0
if. #images_array do.
  images_array=: images_array{~ /: > {.("1) images_array
end.
#images_array
)
prepare_comments=: 3 : 0
if. #comments_array do.
  comments_array=: comments_array{~ /: > {.("1) comments_array
end.
#comments_array
)
prepare_charts=: 3 : 0
if. #charts_array do.
  charts_array=: charts_array{~ /: > {.("1) charts_array
end.
#charts_array
)
store_filtermode=: ''"_
store_autofilterinfo=: ''"_
store_autofilters=: ''"_
store_images=: 3 : 0

recordtype=. 16b00ec                

spid=. {.object_ids
if. 0=#images_array do. '' return. end.

z=. ''
for_i. i.#images_array do.
  'rowcol name xy_offset xy_scale checksum image_id type widthheight'=. i{images_array
  'width height'=. widthheight
  'scale_x scale_y'=. xy_scale
  width=. <.width * (0~:scale_x){1, scale_x
  height=. <.height * (0~:scale_y){1, scale_y
  vertices=. position_object rowcol, xy_offset, width, height

  if. 0=i do.
    dg_length=. 72 + 84* #images_array
    spgr_length=. 48 + 84* #images_array

    dg_length=. dg_length + 120* #charts_array
    spgr_length=. spgr_length + 120* #charts_array

    dg_length=. dg_length + 96* filter_count
    spgr_length=. spgr_length + 96* filter_count

    dg_length=. dg_length + 92* #comments_array
    spgr_length=. spgr_length + 92* #comments_array

    data=. store_mso_dg_container dg_length
    data=. data, store_mso_dg }.object_ids
    data=. data, store_mso_spgr_container spgr_length
    data=. data, store_mso_sp_container 40
    data=. data, store_mso_spgr ''
    data=. data, store_mso_sp 0, spid, 16b0005
    spid=. >:spid
    data=. data, store_mso_sp_container 76
    data=. data, store_mso_sp 75, spid, 16b0a00
    spid=. >:spid
    data=. data, store_mso_opt_image image_id
    data=. data, store_mso_client_anchor 2 ; vertices
    data=. data, store_mso_client_data ''

    z=. z, biffappend data,~ toHeader recordtype, #data

  else.
    data=. store_mso_sp_container 76
    data=. data, store_mso_sp 75, spid, 16b0a00
    spid=. >:spid
    data=. data, store_mso_opt_image image_id
    data=. data, store_mso_client_anchor 2 ; vertices
    data=. data, store_mso_client_data ''

    z=. z, biffappend data,~ toHeader recordtype, #data

  end.

  z=. z, store_obj_image i+1
end.

object_ids=: spid 0}object_ids
z
)
store_charts=: 3 : 0

recordtype=. 16b00ec                

spid=. {.object_ids
num_objects=. #images_array
if. 0=#charts_array do. '' return. end.

z=. ''
for_i. i.#charts_array do.
  'rowcol name xy_offset xy_scale'=. i{charts_array
  'width height'=. 526 319
  'scale_x scale_y'=. xy_scale
  width=. <.width * (0~:scale_x){1, scale_x
  height=. <.height * (0~:scale_y){1, scale_y
  vertices=. position_object rowcol, xy_offset, width, height

  if. (0=i) *. 0=num_objects do.
    dg_length=. 72 + 120* #charts_array
    spgr_length=. 48 + 120* #charts_array

    dg_length=. dg_length + 96* filter_count
    spgr_length=. spgr_length + 96* filter_count

    dg_length=. dg_length + 92* #comments_array
    spgr_length=. spgr_length + 92* #comments_array

    data=. store_mso_dg_container dg_length
    data=. data, store_mso_dg }.object_ids
    data=. data, store_mso_spgr_container spgr_length
    data=. data, store_mso_sp_container 40
    data=. data, store_mso_spgr ''
    data=. data, store_mso_sp 0, spid, 16b0005
    spid=. >:spid
    data=. data, store_mso_sp_container 112
    data=. data, store_mso_sp 201, spid, 16b0a00
    spid=. >:spid
    data=. data, store_mso_opt_chart ''
    data=. data, store_mso_client_anchor 0 ; vertices
    data=. data, store_mso_client_data ''

    z=. z, biffappend data,~ toHeader recordtype, #data

  else.
    data=. store_mso_sp_container 112
    data=. data, store_mso_sp 201, spid, 16b0a00
    spid=. >:spid
    data=. data, store_mso_opt_chart ''
    data=. data, store_mso_client_anchor 0 ; vertices
    data=. data, store_mso_client_data ''

    z=. z, biffappend data,~ toHeader recordtype, #data

  end.

  z=. z, store_obj_chart num_objects+i+1
  store_chart_binary name
end.
formula=. '=', sheetname, '!A1'
object_ids=: spid 0}object_ids
z
)
store_chart_binary=: 3 : 0

if. _1-: z=. fread y do. 'Couldn''t open file in add_chart_ext' 13!:8 (3) end.
z=. biffappend z
)
store_filters=: 3 : 0

recordtype=. 16b00ec                

spid=. {.object_ids
num_objects=. (#images_array) + #charts_array
if. 0=filter_count do. '' return. end.

'row1 row2 col1 col2'=. filter_area

z=. ''
for_i. i.filter_count do.

  vertices=. (col1 +i), 0, row1, 0, (col1 +i +1), 0, (row1 +1), 0

  if. (0=i) *. 0=num_objects do.
    dg_length=. 72 + 96* filter_count
    spgr_length=. 48 + 96* filter_count

    dg_length=. dg_length + 92* #comments_array
    spgr_length=. spgr_length + 92* #comments_array

    data=. store_mso_dg_container dg_length
    data=. data, store_mso_dg }.object_ids
    data=. data, store_mso_spgr_container spgr_length
    data=. data, store_mso_sp_container 40
    data=. data, store_mso_spgr ''
    data=. data, store_mso_sp 0, spid, 16b0005
    spid=. >:spid
    data=. data, store_mso_sp_container 88
    data=. data, store_mso_sp 201, spid, 16b0a00
    spid=. >:spid
    data=. data, store_mso_opt_filter ''
    data=. data, store_mso_client_anchor 1 ; vertices
    data=. data, store_mso_client_data ''

    z=. z, biffappend data,~ toHeader recordtype, #data

  else.
    data=. store_mso_sp_container 88
    data=. data, store_mso_sp 201, spid, 16b0a00
    spid=. >:spid
    data=. data, store_mso_opt_filter ''
    data=. data, store_mso_client_anchor 1 ; vertices
    data=. data, store_mso_client_data ''

    z=. z, biffappend data,~ toHeader recordtype, #data

  end.

  z=. z, store_obj_filter (num_objects+i+1), col1 +i
end.
formula=. '=', sheetname, '!A1'
z=. z, store_formula formula

object_ids=: spid 0}object_ids
z
)
store_comments=: 3 : 0

recordtype=. 16b00ec                

spid=. {.object_ids
num_objects=. (#images_array) + filter_count + #charts_array
if. 0=#comments_array do. '' return. end.

z=. ''
for_i. i.#comments_array do.
  'rowcol str author visible color vertices'=. i{comments_array
  formats=. 0 5 ,: (#str), 0
  if. (0=i) *. 0=num_objects do.
    dg_length=. 72 + 92* #comments_array
    spgr_length=. 48 + 92* #comments_array

    data=. store_mso_dg_container dg_length
    data=. data, store_mso_dg }.object_ids
    data=. data, store_mso_spgr_container spgr_length
    data=. data, store_mso_sp_container 40
    data=. data, store_mso_spgr ''
    data=. data, store_mso_sp 0, spid, 16b0005
    spid=. >:spid
    data=. data, store_mso_sp_container 84
    data=. data, store_mso_sp 202, spid, 16b0a00
    spid=. >:spid
    data=. data, store_mso_opt_comment visible
    data=. data, store_mso_client_anchor 3 ; vertices
    data=. data, store_mso_client_data ''

    z=. z, biffappend data,~ toHeader recordtype, #data

  else.
    data=. store_mso_sp_container 84
    data=. data, store_mso_sp 202, spid, 16b0a00
    spid=. >:spid
    data=. data, store_mso_opt_comment visible
    data=. data, store_mso_client_anchor 3 ; vertices
    data=. data, store_mso_client_data ''

    z=. z, biffappend data,~ toHeader recordtype, #data

  end.

  z=. z, store_obj_comment num_objects+i+1
  z=. z, store_mso_drawing_text_box ''
  z=. z, store_txo #str
  z=. z, store_txo_continue_1 str
  z=. z, store_txo_continue_2 formats
end.
for_i. i.#comments_array do.
  'rowcol str author visible color vertices'=. i{comments_array
  z=. z, store_note rowcol ; (num_objects+i+1) ; author ; visible
end.
z
)
store_mso_dg_container=: 3 : 0

type=. 16bf002
version=. 15
'' add_mso_generic type, version, 0, y
)
store_mso_dg=: 3 : 0

type=. 16bf008
version=. 0
'instance num_shapes max_spid'=. y
data=. toDWORD0 num_shapes, max_spid
data add_mso_generic type, version, instance, #data
)
store_mso_spgr_container=: 3 : 0

type=. 16bf003
version=. 15
'' add_mso_generic type, version, 0, y
)
store_mso_sp_container=: 3 : 0

type=. 16bf004
version=. 15
'' add_mso_generic type, version, 0, y
)
store_mso_spgr=: 3 : 0

type=. 16bf009
version=. 1
data=. toDWORD0 0 0 0 0
data add_mso_generic type, version, 0, #data
)
store_mso_sp=: 3 : 0

type=. 16bf00a
version=. 2
'instance spid options'=. y
data=. toDWORD0 spid, options
data add_mso_generic type, version, instance, #data
)
store_mso_opt=: 3 : 0

type=. 16bf00b
version=. 3
instance=. 9

length=. 54   

'spid visible color'=. 3{.y, 16b50
if. _1~:visible do.
  hidden=. (0=visible) { 16b0000 16b0002
else.
  hidden=. (0=comments_visible) { 16b0000 16b0002
end.

data=. toDWORD0 spid
data=. data, toBYTE 16b00, 16b00, 16bbf, 16b00, 16b08, 16b00, 16b08, 16b00, 16b58, 16b01, 16b00, 16b00, 16b00, 16b00, 16b81, 16b01
data=. data, toBYTE color
data=. data, toBYTE 16b00, 16b00, 16b08, 16b83, 16b01, 16b50, 16b00, 16b00, 16b08, 16bbf, 16b01, 16b10, 16b00, 16b11, 16b00, 16b01, 16b02, 16b00, 16b00, 16b00, 16b00, 16b3f, 16b02, 16b03, 16b00, 16b03, 16b00, 16bbf, 16b03
data=. data, toWORD0 hidden
data=. data, toBYTE 16b0a, 16b00

assert. length = #data
data add_mso_generic type, version, instance, #data
)
store_mso_opt_image=: 3 : 0

type=. 16bf00b
version=. 3
instance=. 3

length=. 6 * instance   

spid=. y

data=. toWORD0 16b4104                 
data=. data, toDWORD0 spid
data=. data, toWORD0 16b01bf           
data=. data, toDWORD0 16b00010000
data=. data, toWORD0 16b03bf           
data=. data, toDWORD0 16b00080000

assert. length = #data
data add_mso_generic type, version, instance, #data
)
store_mso_opt_chart=: 3 : 0

type=. 16bf00b
version=. 3
instance=. 9

length=. 6 * instance   

data=. toWORD0 16b007f                  
data=. data, toDWORD0 16b01040104

data=. data, toWORD0 16b00bf            
data=. data, toDWORD0 16b00080008

data=. data, toWORD0 16b0181            
data=. data, toDWORD0 16b0800004e

data=. data, toWORD0 16b0183            
data=. data, toDWORD0 16b0800004d

data=. data, toWORD0 16b01bf            
data=. data, toDWORD0 16b00110010

data=. data, toWORD0 16b01c0            
data=. data, toDWORD0 16b0800004d

data=. data, toWORD0 16b01ff            
data=. data, toDWORD0 16b00080008

data=. data, toWORD0 16b023f            
data=. data, toDWORD0 16b00020000

data=. data, toWORD0 16b03bf            
data=. data, toDWORD0 16b00080000

assert. length = #data
data add_mso_generic type, version, instance, #data
)
store_mso_opt_filter=: 3 : 0

type=. 16bf00b
version=. 3
instance=. 5

length=. 6 * instance   

data=. toWORD0 16b007f                  
data=. data, toDWORD0 16b01040104

data=. data, toWORD0 16b00bf            
data=. data, toDWORD0 16b00080008

data=. data, toWORD0 16b01bf            
data=. data, toDWORD0 16b00010000

data=. data, toWORD0 16b01ff            
data=. data, toDWORD0 16b00080000

data=. data, toWORD0 16b03bf            
data=. data, toDWORD0 16b000a0000

assert. length = #data
data add_mso_generic type, version, instance, #data
)
store_mso_opt_comment=: 3 : 0

type=. 16bf00b
version=. 3
instance=. 3

length=. 6 * instance   

visible=. y
if. _1~:visible do.
  hidden=. (0=visible) { 16b0000 16b0002
else.
  hidden=. (0=comments_visible) { 16b0000 16b0002
end.

data=. toWORD0 16b0080            
data=. data, toDWORD0 0
data=. data, toWORD0 16b03be            
data=. data, toWORD0 hidden, 16b0002
data=. data, toWORD0 16b03bf            
data=. data, toWORD0 16b0002, 16b0002

assert. length = #data
data add_mso_generic type, version, instance, #data
)
store_mso_client_anchor=: 3 : 0

type=. 16bf010
version=. 0

'flag vertices'=. y
'col_start x1 row_start y1 col_end x2 row_end y2'=. vertices
data=. toWORD0 flag, col_start, x1, row_start, y1, col_end, x2, row_end, y2

data add_mso_generic type, version, 0, #data
)
store_mso_client_data=: 3 : 0

type=. 16bf011
version=. 0
'' add_mso_generic type, version, 0, 0
)
store_obj_comment=: 3 : 0

recordtype=. 16b005d        
length=. 16b0034            

obj_id=. y                  
obj_type=. 16b0019          

options=. 16b4011
reserved=. 16b0000
sub_recordtype=. 16b0015             
sub_data=. toWORD0 obj_type, obj_id, options
sub_data=. sub_data, toDWORD0 reserved, reserved, reserved
assert. 16b0012 = #sub_data
data=. sub_data,~ toHeader sub_recordtype, #sub_data
sub_recordtype=. 16b000d             
sub_data=. toDWORD0 5#reserved
sub_data=. sub_data, toWORD0 reserved
assert. 16b0016 = #sub_data
data=. data, sub_data,~ toHeader sub_recordtype, #sub_data
sub_recordtype=. 16b0000                 
data=. data, toHeader sub_recordtype, 0

assert. length = #data
z=. biffappend data,~ toHeader recordtype, length
)
store_obj_image=: 3 : 0

recordtype=. 16b005d        
length=. 16b0026            

obj_id=. y                  
obj_type=. 16b0008          

options=. 16b6011
reserved=. 16b0000
sub_recordtype=. 16b0015             
sub_data=. toWORD0 obj_type, obj_id, options
sub_data=. sub_data, toDWORD0 reserved, reserved, reserved
assert. 16b0012 = #sub_data
data=. sub_data,~ toHeader sub_recordtype, #sub_data
sub_recordtype=. 16b0007             
sub_data=. toWORD0 16bffff
assert. 16b0002 = #sub_data
data=. data, sub_data,~ toHeader sub_recordtype, #sub_data
sub_recordtype=. 16b0008                  
sub_data=. toWORD0 16b0001
assert. 16b0002 = #sub_data
data=. data, sub_data,~ toHeader sub_recordtype, #sub_data
sub_recordtype=. 16b0000                 
data=. data, toHeader sub_recordtype, 0

assert. length = #data
z=. biffappend data,~ toHeader recordtype, length
)
store_obj_chart=: 3 : 0

recordtype=. 16b005d        
length=. 16b001a            

obj_id=. y                  
obj_type=. 16b0005          

options=. 16b6011
reserved=. 16b0000
sub_recordtype=. 16b0015             
sub_data=. toWORD0 obj_type, obj_id, options
sub_data=. sub_data, toDWORD0 reserved, reserved, reserved
assert. 16b0012 = #sub_data
data=. sub_data,~ toHeader sub_recordtype, #sub_data
sub_recordtype=. 16b0000                 
data=. data, toHeader sub_recordtype, 0

assert. length = #data
z=. biffappend data,~ toHeader recordtype, length
)
store_obj_filter=: 3 : 0

recordtype=. 16b005d        
length=. 16b0046        

'obj_id col'=. y        
obj_type=. 16b0014        

options=. 16b2101
reserved=. 16b0000
sub_recordtype=. 16b0015             
sub_data=. toWORD0 obj_type, obj_id, options
sub_data=. sub_data, toDWORD0 reserved, reserved, reserved
assert. 16b0012 = #sub_data
data=. sub_data,~ toHeader sub_recordtype, #sub_data
sub_recordtype=. 16b000c             
sub_data=. toBYTE 16b00, 16b00, 16b00, 16b00, 16b00, 16b00, 16b00, 16b00, 16b64, 16b00, 16b01, 16b00, 16b0a, 16b00, 16b00, 16b00, 16b10, 16b00, 16b01, 16b00
assert. 16b0014 = #sub_data
data=. data, sub_data,~ toHeader sub_recordtype, #sub_data
sub_recordtype=. 16b0013             
if. col e. >{.("1) filter_cols do.
  sub_data=. toBYTE 16b00, 16b00, 16b00, 16b00, 16b01, 16b00, 16b01, 16b03, 16b00, 16b00, 16b0a, 16b00, 16b08, 16b00, 16b57, 16b00
else.
  sub_data=. toBYTE 16b00, 16b00, 16b00, 16b00, 16b01, 16b00, 16b01, 16b03, 16b00, 16b00, 16b02, 16b00, 16b08, 16b00, 16b57, 16b00
end.

assert. 16b1fee = #sub_data
data=. data, sub_data,~ toHeader sub_recordtype, #sub_data
sub_recordtype=. 16b0000                 
data=. data, toHeader sub_recordtype, 0

assert. length = #data
z=. biffappend data,~ toHeader recordtype, length
)
store_mso_drawing_text_box=: 3 : 0

recordtype=. 16b00ec                
length=. 16b0008                

data=. store_mso_client_text_box ''

assert. length = #data
z=. biffappend data,~ toHeader recordtype, length
)
store_mso_client_text_box=: 3 : 0

type=. 16bf00d
version=. 0

'' add_mso_generic type, version, 0, 0
)
store_txo=: 3 : 0

recordtype=. 16b01b6               
length=. 16b0012                   
'string_len format_len rotation'=. 3{.y
format_len=. (2>#y){format_len, 16          
grbit=. 16b0212                  
reserved=. 16b0000                  
data=. toWORD0 grbit, rotation
data=. data, toDWORD0 reserved
data=. data, toWORD0 reserved, string_len, format_len
data=. data, toDWORD0 reserved

assert. length = #data
z=. biffappend data,~ toHeader recordtype, length
)
store_txo_continue_1=: 3 : 0

recordtype=. 16b003c         
string=. y                   
if. encoding=. 2~: 3!:0 string do.  
  string=. toucode0 string
end.
limit=. RECORDLEN-6  
z=. ''
while. limit<#string do.
  string=. limit}.string [ tmp_str=. limit{.string
  z=. z, biffappend data,~ toHeader recordtype, #data=. tmp_str,~ toBYTE encoding
end.
z=. z, biffappend data,~ toHeader recordtype, #data=. string,~ toBYTE encoding
)
store_txo_continue_2=: 3 : 0

recordtype=. 16b003c                   
formats=. y                   
data=. ''
for_aref. formats do.
  data=. data, toWORD0 aref
  data=. data, toDWORD0 16b0
end.
z=. biffappend data,~ toHeader recordtype, #data
)
store_note=: 3 : 0

recordtype=. 16b001c                   
length=. 16b000c                   
'rowcol obj_id author visible'=. y
author=. (''-:author){:: author ; comments_author
if. _1~:visible do.
  hidden=. (0=visible) { 16b0002 16b0000
else.
  hidden=. (0=comments_visible) { 16b0002 16b0000
end.
num_chars=. #author
author=. author, {.a.
data=. toWORD0 rowcol, hidden, obj_id, num_chars
if. 2= (3!:0) author do.
  data=. data, toBYTE 0
  z=. biffappend (data, author),~ toHeader recordtype, (#data) + #author
else.
  data=. data, toBYTE 1
  z=. biffappend (data, author),~ toHeader recordtype, (#data) + #author=. toucode0 author
end.
z
)

defcommentparams=: 0 : 0
author ''
color _1
start_cell _1 _1
start_col _1
start_row _1
visible _1
width 128
height 74
x_offset _1
x_scale 1
y_offset _1
y_scale 1
)
comment_params=: 3 : 0

'rowcol string'=. 2{.y
nm=. (<'params_'),&.> {.("1) nmv=. (_2]\ 2}.y),~ ;:;._2 defcommentparams
(, (1&u: :: ])&.> nm)=. ".&.>{:("1) nmv
params_width=. (0<:params_width){ params_width, 128
params_height=. (0<:params_height){ params_height, 74
max_len=. 32767
if. max_len < #string do. string=. max_len{.string end.
color=. params_color
color=. color
color=. (color=16b7fff){ color, 16b50   
params_color=. color
if. _1 _1-.@-:params_start_cell do.
  if. 2 131072 e.~ 3!:0 params_start_cell do. params_start_cell=. A1toRC params_start_cell end.
  rowcol=. 'params_start_row params_start_col'=. params_start_cell
end.

'row col'=. rowcol
if. _1=params_start_row do.
  if. 0=col do. params_start_row=. 0
  elseif. 65533=row do. params_start_row=. 65529
  elseif. 65534=row do. params_start_row=. 65530
  elseif. 65535=row do. params_start_row=. 65531
  elseif. do. params_start_row=. row -1
  end.
end.

if. _1=params_y_offset do.
  if. 0=row do. params_y_offset=. 2
  elseif. 65533=row do. params_y_offset=. 4
  elseif. 65534=row do. params_y_offset=. 4
  elseif. 65535=row do. params_y_offset=. 2
  elseif. do. params_y_offset=. 7
  end.
end.

if. _1=params_start_col do.
  if. 253=col do. params_start_col=. 250
  elseif. 254=col do. params_start_col=. 251
  elseif. 255=col do. params_start_col=. 252
  elseif. do. params_start_col=. col +1
  end.
end.

if. _1=params_x_offset do.
  if. 253=col do. params_x_offset=. 49
  elseif. 254=col do. params_x_offset=. 49
  elseif. 255=col do. params_x_offset=. 49
  elseif. do. params_x_offset=. 15
  end.
end.
if. params_x_scale do.
  params_width=. <. params_width * params_x_scale
end.

if. params_y_scale do.
  params_height=. <. params_height * params_y_scale
end.
vertices=. position_object params_start_col, params_start_row, params_x_offset, params_y_offset, params_width, params_height

(row, col) ; string ; params_author ; params_visible ; params_color ; vertices
)
coclass 'biffbook'
coinsert 'biff'
embedchart=: 4 : 0
l=. sheeti{sheet
x embedchart__l y
)
insertimage=: 4 : 0
l=. sheeti{sheet
x insertimage__l y
)
writecomment=: 4 : 0
l=. sheeti{sheet
x writecomment__l y
)
add_mso_drawing_group=: 3 : 0

if. 0=mso_size do. '' return. end.

recordtype=. 16b00eb               

data=. store_mso_dgg_container ''
data=. data, store_mso_dgg mso_clusters
data=. data, store_mso_bstore_container ''
for_im. images_data do. data=. data, store_mso_images im end.
data=. data, store_mso_opt ''
data=. data, store_mso_split_menu_colors ''

header=. toHeader recordtype, #data

z=. add_mso_drawing_group_continue header, data
z
)
add_mso_drawing_group_continue=: 3 : 0

data=. y
limit=. RECORDLEN  
mso_group=. 16b00eb    
continue=. 16b003c    
block_count=. 1
if. (#data) <: limit do.
  z=. 1&biffappend data
  return.
end.

z=. ''
data=. (limit+4)}.data [ tmp=. (limit+4){.data
tmp=. (toWORD0 limit) 2 3}tmp
z=. z, 1&biffappend tmp
while. (#data) > limit do.
  if. (block_count = 1) do.
    header=. toHeader mso_group, limit
    block_count=. >:block_count
  else.
    header=. toHeader continue, limit
  end.

  data=. limit}.data [ tmp=. limit{.data
  z=. z, 1&biffappend header, tmp
end.
z=. z, 1&biffappend data,~ toHeader continue, #data
z
)
store_mso_dgg_container=: 3 : 0

type=. 16bf000
version=. 15
data=. ''
length=. mso_size -12    

data add_mso_generic type, version, 0, length
)
store_mso_dgg=: 3 : 0

type=. 16bf006
version=. 0

'max_spid num_clusters shapes_saved drawings_saved clusters'=. y

data=. toDWORD0 max_spid, num_clusters, shapes_saved, drawings_saved

data=. data, toDWORD0 ,clusters

data add_mso_generic type, version, 0, #data  
)
store_mso_bstore_container=: 3 : 0

if. 0=images_size do. '' return. end.

type=. 16bf001
version=. 15
instance=. #images_data    
data=. ''
length=. images_size +8 *instance

data add_mso_generic type, version, instance, length
)
store_mso_images=: 3 : 0

'ref_count image_type image size checksum1 checksum2'=. y

blip_store_entry=. store_mso_blip_store_entry ref_count ; image_type ; size ; checksum1

blip=. store_mso_blip image_type ; image ; size ; checksum1 ; checksum2

blip_store_entry , blip
)
store_mso_blip_store_entry=: 3 : 0

'ref_count image_type size checksum1'=. y

type=. 16bf007
version=. 2
instance=. image_type
length=. size +61
data=. toBYTE image_type   
data=. data, toBYTE image_type   
data=. data, checksum1           
data=. data, toWORD0 16bff          
data=. data, toDWORD0 size +25     
data=. data, toDWORD0 ref_count    
data=. data, toDWORD0 16b00000000    
data=. data, toBYTE 16b00          
data=. data, toBYTE 16b00          
data=. data, toBYTE 16b00          
data=. data, toBYTE 16b00          
data add_mso_generic type, version, instance, length
)
store_mso_blip=: 3 : 0

'image_type image_data size checksum1 checksum2'=. y

if. image_type = 5 do. instance=. 16b046a end.   
if. image_type = 6 do. instance=. 16b06e0 end.   
if. image_type = 7 do. instance=. 16b07a9 end.   
if. ( image_type = 7) do.
  checksum1=. checksum2, checksum1
end.

type=. 16bf018 + image_type
version=. 16b0000
length=. size +17
data=. checksum1          
data=. data, toBYTE 16bff       
data=. data, image_data            

assert. length=#data
data add_mso_generic type, version, instance, length
)
store_mso_opt=: 3 : 0

type=. 16bf00b
version=. 3

data=. toBYTE 16bfb, 16b00, 16b80, 16b00, 16b80, 16b00, 16b18, 16b10, 16b90, 16b00, 16b00, 16b80, 16b0c, 16b10, 16b04, 16b00, 16b00, 16b80
data add_mso_generic type, version, 3, #data
)
store_mso_split_menu_colors=: 3 : 0

type=. 16bf11e
version=. 0

data=. toBYTE 16bd0, 16b00, 16b00, 16b80, 16bc0, 16b00, 16b00, 16b80, 16b71, 16b00, 16b00, 16b80, 16b7f, 16b00, 16b00, 16b01
data add_mso_generic type, version, 4, #data
)
calc_sheet_offsets=: 3 : 0

BOF=. 12
EOF=. 4
offset=. datasize
offset=. offset + calculate_shared_string_sizes ''
offset=. offset + calculate_extern_sizes ''
offset=. offset + mso_size + 4 * <.(mso_size -1) % limit

for_wsheet. sheet do.
  offset=. offset + BOF + #name__wsheet   
end.

offset=. offset + EOF

for_wsheet. sheet do.
  offset=. offset + datasize__wsheet
end.

biffsize=: offset
)
calc_mso_sizes=: 3 : 0

msoSize=. 0     
start_spid=. 1024  
max_spid=. 1024  
num_clusters=. 1     
shapes_saved=. 0     
drawings_saved=. 0     
clusters=. 0 2$''

process_images ''
msoSize=. msoSize + (0~:#images_data){0 8
for_wsheet. sheet do.
  if. 'biffsheet' -.@-: >@coinstance wsheet do. continue. end.

  if. 0= num_shapes=. (#images_array__wsheet) + filter_count__wsheet + (prepare_comments__wsheet '') + (prepare_charts__wsheet '') do. continue. end.
  shapes_saved=. shapes_saved + num_shapes=. 1 + num_shapes
  msoSize=. msoSize + image_mso_size__wsheet
  drawings_saved=. >:drawings_saved
  max_spid=. 1024 * 1 + <.(max_spid -1)%1024
  start_spid=. max_spid
  max_spid=. max_spid + num_shapes
  i=. num_shapes
  while. i > 0 do.
    num_clusters=. num_clusters + 1
    msoSize=. msoSize + 8
    size=. (i <: 1024) { 1024, i
    clusters=. clusters, drawings_saved, size
    i=. i - 1024
  end.
  object_ids__wsheet=: start_spid, drawings_saved, num_shapes, max_spid -1
end.
mso_size=: msoSize + (0~:msoSize){0 86  
mso_clusters=: max_spid ; num_clusters ; shapes_saved ; drawings_saved ; clusters
)
process_images=: 3 : 0

images_seen=. 0 4$''
imageData=. 0 0$''
image_id=. 1
imageSize=. 0

for_wsheet. sheet do.
  if. 'biffsheet' -.@-: >@coinstance wsheet do. continue. end.
  if. 0=prepare_images__wsheet '' do. continue. end.

  image_msoSize=. 0
  imagesArray=. 0 0$''
  for_imageref. images_array__wsheet do.
    if. (16#{.a.)-: 4{::imageref do.
      filename=. 1{imageref
      if. _1-: data=. fread filename do.
        'Couldn''t import file' 13!:8 (3)
      end.
    elseif. (16#{:a.)-: 4{::imageref do.
      data=. 1{:: imageref
    elseif. do.
      'unhandled exception' 13!:8 (3)
    end.

    if. ({.("1)images_seen) -.@e.~ <ck=. image_checksum data do.
      size=. #data
      checksum1=. image_checksum data
      checksum2=. checksum1
      ref_count=. 1
      if. (}.4{.data) -: 'PNG' do.
        'type width height'=. process_png data

      elseif. (2{.data) -: 'BM' do.
        'type width height'=. process_bmp data
        data=. 14}.data
        checksum2=. image_checksum data
        size=. size + 2

      elseif. do.
        if. 0 0 0 -: 'type width height'=. process_jfif data do.
          'Unsupported image format for file' 13!:8 (3)
        end.
      end.
      imagesArray=. imagesArray, (_4}.imageref), checksum1 ; image_id ; type ; width, height
      imageData=. imageData, ref_count ; type ; data ; size ; checksum1 ; checksum2
      imageSize=. imageSize + size +61  
      image_msoSize=. image_msoSize + size +69  

      images_seen=. images_seen, checksum1 ; image_id ; type ; width, height
      image_id=. >:image_id
    else.
      index=. ({.("1)images_seen) i. <ck
      imageData=. (>:&.> (<index, 0){imageData) (<index, 0)}imageData
      imagesArray=. imagesArray, (_4}.imageref), index{images_seen
    end.
  end.
  assert. (#imagesArray) = #images_array__wsheet
  images_array__wsheet=: imagesArray
  image_mso_size__wsheet=: image_msoSize
end.
images_size=: imageSize
images_data=: imageData  
)
image_checksum=: 3 : 0

md4_biffmd4_ y
)
process_jfif=: 3 : 0

type=. 5  

if. 1 e. a=. (16bff 16bd8 16bff{a.) E. y do.
  p=. 2+ a i. 1
  try.
    while. do.
      while. -. ((16bff{a.)=p{y) *. (16bff{a.)~:(p+1){y do.
        p=. p+1
      end.
      p=. p+1
      if. (p{y) e. 16bc0 16bc1 16bc2 16bc3 16bc5 16bc6 16bc7 16bc9 16bca 16bcb 16bcd 16bce 16bcf{a. do.
        break.
      end.
      p=. p + (a.i.(p+2){y)+ 256 * a.i.(p+1){y
    end.
    h=. (a.i.(p+5){y)+ 256 * a.i.(p+4){y
    w=. (a.i.(p+7){y)+ 256 * a.i.(p+6){y
    depth=. 8 * a.i.(p+8){y
    type, w, h
  catch.
    0 0 0
  end.
else.
  0 0 0
end.
)
process_png=: 3 : 0

type=. 6  
'width height'=. fromDWORD1 (16+i.8){y

type, width, height
)
process_bmp=: 3 : 0

data=. y
type=. 7 
if. (#data) <: 16b36 do.
  'file: doesn''t contain enough data.' 13!:8 (3)
end.
'width height'=. fromDWORD0 (18+i.8){data

if. (width > 16bffff) do.
  'file: largest image width width supported is 65k.' 13!:8 (3)
end.

if. (height > 16bffff) do.
  'file: largest image height supported is 65k.' 13!:8 (3)
end.
'planes bitcount'=. fromWORD0 (26+i.4){data
if. (planes ~: 1) do.
  'file: only 1 plane supported in bitmap image.' 13!:8 (3)
end.
compression=. fromDWORD0 (30+i.4){data

if. (compression ~: 0) do.
  'file: compression not supported in bitmap image.' 13!:8 (3)
end.

type, width, height
)
