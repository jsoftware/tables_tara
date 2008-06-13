NB. =========================================================
NB. Cover for writing one or more arrays to sheets in an Excel workbook
NB. written by Ric Sherlock

coxclass 'biffwrite'
coxtend 'biffbook'

NB. ---------------------------------------------------------
NB.*writeexcelsheets v Write arrays to sheets of Excel workbook
NB. returns: numeric 1 if successful.
NB. form: ([<sheetname(s)>],.<sheetcontent(s)>) writesheets <filename>
NB. y is: literal filename of workbook to create
NB. x is: One of;
NB.     * literal (data to write in topleft cell of Sheet1)
NB.     * numeric matrix (data to write to Sheet1)
NB.     * boxed (numeric/literal/mixed) matrix (data to write to Sheet1)
NB.     * 2-item/column vector/matrix,
NB.          Sheetnames in 1st item/col (literal)
NB.          Associated data formats (above) for sheetnames
NB. if <sheetname(s)> not given then defaults used 
writeexcelsheets=: 4 : 0
  y=. y. [ x=. x.
  if. 0=#x do. empty'' return. end. NB. if empty xarg then return.
  shts=. makexarg x
  shtnme=. ((0 < #) {:: 'Sheet1'&;) (<0 0) {:: shts
  bi=. ('Arial' ; 220 ; shtnme) coxnew 'biffbook'
  shtdat=. (<0 1){:: shts
  bi writeshtdat shtdat NB. write data for first worksheet
  shts=. }.shts         NB. drop first sheet from list
  bi addsheets"1 shts NB. add and write to rest of sheets
  save__bi y
  destroy__bi ''
)

NB. ---------------------------------------------------------
NB. utils
mfv1=: ,:^:(#&$ = 1:)       NB. makes 1-row matrix from vector
mfva=: ,:^:([: 2&> #@$)^:_  NB. makes a matrix from an atom or vector
ischar=: 3!:0 e. 2 131072"_
firstones=: > 0: , }:
lastones=: > 0: ,~ }.
NB. indices=: $ #: I.@,   NB. get row,.col indices of 1s in matrix
indices=: 4$.$.       NB. get row,.col indices of 1s in matrix
isrowblks=: >/@(+/^:(#&$)"_1) NB. are blocks row oriented

NB. isdata v Verb decides on type of x argument to writesheets
NB. returns 0 for array of names and data, 1 for single data matrix.
isdata=: 3 : 0
  y=. y.
  if. 1< lvls=. L. y do. 0 return. NB. if 1<L.x must be multiple sheets
  elseif. 0= lvls do. 1 return. NB. if not boxed then must be data for single sheet
  elseif. 2< {:$ y do. 1 return. NB. if more than 2 items/cols then data for one sheet
  elseif. -. *./2= 3!:0 &> {."1 y do. 1 return. NB. if not all 1st col contents are literal then must be one sheet
  elseif. 1< #@$ &> {:"1 y do. 0 return. NB. if any last col contents have rank greater than 1 then must be multiple sheets
  elseif. do. 1 NB. else assume that boxed data for one sheet
  end.
)

NB. makexarg v Ensures that xarg to writesheets has right form.
NB. returns: 2-item/column vector/array.
NB.       {."1 are sheetnames, {:"1 are boxed rank-2 arrays of sheetdata
makexarg=: 3 : 0
  y=. y.
  if. isdata y do.
    if. 2= 3!:0 y do. y=. <y end.
    y=. a:,. <mfva y
  else.
    if. 2>{:$ y do. NB. if only 1 item/col add empty column of sheetnames
      y=. a:,.y
    end.
    if. #idx=. (I. b=. ischar &>{:"1 y) do.  NB. sheets with unboxed string data
      upd=. ({."1 ,. <&.>@({:"1)) b#y NB. boxed versions
      y=. upd idx }y
    end.
    if. #idx=. (I. b=.2> #@$&>{:"1 y) do.  NB. sheets with unboxed string data
      upd=. ({."1 ,. mfva&.>@({:"1)) b#y NB. boxed versions
      y=. upd idx }y
    end.
  end.
  mfv1 y
)

NB. addsheets v Creates new sheet and writes data to it
NB. form: <wkbklocale> addSheets <sheetname>;<array>
addsheets=: 4 : 0
  y=. y. [ x=. x.
  'shtnme shtdat'=. y
  addsheet__x shtnme
  x writeshtdat shtdat
)

NB. writeshtdat v Writes array to current worksheet
NB. form: <wkbklocale> writeShtdat <array>
NB. Writes blocks of string data and numeric data separately
NB. Only 1d blocks at present. to minimise the number of calls it would be good to
NB. find rectangular blocks of same type and write them
NB. using (<corner>,:<shape>) u;.0 <array>
writeshtdat=: 4 : 0
  y=. y. [ x=. x.
  if. 0=L.y do.
    writenumber__x 0 0;y
  else.
    as=. ischar &> y
    blks=. blocksx as
    tls=. {.0 2|: blks
    dat=: blks <;.0 y NB. blocks of char
    writestring__x"1 (<"1 tls),.dat
    blks=. blocksx -.as
    tls=. {.0 2|: blks
    dat=. blks ([:<>);.0 y  NB. blocks of non-char
    writenumber__x"1 (<"1 tls),.dat
  end.
)

NB. ---------------------------------------------------------
NB. tests
Note 'testargs'
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
)

Note 'tests for writeexcelsheets'
x1 writexlsheets jpath '~temp/tarawsht1.xls'
x2 writexlsheets jpath '~temp/tarawsht2.xls'
x3 writexlsheets jpath '~temp/tarawsht3.xls'
x4 writexlsheets jpath '~temp/tarawsht4.xls'
x5 writexlsheets jpath '~temp/tarawsht5.xls'
x6 writexlsheets jpath '~temp/tarawsht6.xls'
x7 writexlsheets jpath '~temp/tarawsht7.xls'
x8 writexlsheets jpath '~temp/tarawsht8.xls'
x9 writexlsheets jpath '~temp/tarawsht9.xls'
x10 writexlsheets jpath '~temp/tarawsht10.xls'
x11 writexlsheets jpath '~temp/tarawsht11.xls'
x12 writexlsheets jpath '~temp/tarawsht12.xls'
x13 writexlsheets jpath '~temp/tarawsht13.xls'
x14 writexlsheets jpath '~temp/tarawsht14.xls'
x15 writexlsheets jpath '~temp/tarawsht15.xls'
x16 writexlsheets jpath '~temp/tarawsht16.xls'
x17 writexlsheets jpath '~temp/tarawsht17.xls'
x18 writexlsheets jpath '~temp/tarawsht18.xls'
x19 writexlsheets jpath '~temp/tarawsht19.xls'
)

NB. ---------------------------------------------------------
NB. working 1d solutions for creating blocks.

NB. explicit solution (chooses best block orientation)
blocksx=: 3 : 0
  y=. y.
  fo=. (firstones ,: firstones"1) y
  if. isr=. isrowblks fo do. NB. row-oriented
    tl=. indices isr { fo
    br=. indices lastones"1 y
  else. NB. column-oriented
    tl=. |."1 indices |: isr { fo
    br=. |."1 indices |: lastones y
  end.
  tl ,:"1 >:br-tl
)

NB. ---------------------------------------------------------
Note 'tacit solution'
NB. (works best for row-oriented blocks)
tls=: [: indices firstones"1 NB. topleft index of blocks of 1s
brs=: [: indices lastones"1  NB. bottomright index of blocks of 1s

shapes=: [: >: brs - tls  NB. shapes of blocks of 1s
blocks=: tls ,:"1 shapes  NB. blocks of 1s
)

Note 'testing'
tls tst1          NB. list of topleft of blocks of 1s
tls -.tst1        NB. list of topleft of blocks of 0s

(blocks ischar &> tsta) <;.0 tsta  NB. blocks of char (you will need to create a tsta to run this)
(blocksx -.ischar &> tsta) <;.0 tsta  NB. blocks of non-char
(blocksx -.ischar &> tsta) ([:<>);.0 tsta  NB. blocks of non-char
)

NB. =========================================================
NB. Attempt 2d version for creating blocks
Note '2d version of blocks'
tl=: (firstones *. firstones"1)  NB. topright
br=: (lastones *. lastones"1)    NB. bottomleft

tsta=: readexcel jpath '~temp/tararead.xls'
tst1=: ischar every tsta

('tst1';'firstones';'lastones';'firstones"1';'lastones"1'),:(;firstones;lastones;firstones"1;lastones"1) tst1
indices tl tst1
indices br tst1
)

NB. =========================================================
NB.  publish in z locale
writexlsheets_z_=: writeexcelsheets_biffwrite_