NB. ---------------------------------------------------------
NB. package for biff format
coxclass 'biff'
coxtend 'oleutlfcn'
shortdatefmt=: 'dd/mm/yyyy'
RECORDLEN=: 8224   NB. BIFF5: 2080 bytes, BIFF8: 8224 bytes
NB. Excel version BIFF version Document type File type
NB. Excel 2.1 BIFF2 Worksheet Stream
NB. Excel 3.0 BIFF3 Worksheet Stream
NB. Excel 4.0 BIFF4S Worksheet Stream
NB. Excel 4.0 BIFF4W Workbook Stream
NB. Excel 5.0 BIFF5 Workbook Compound Document
NB. Excel 7.0 (Excel 95) BIFF7 Workbook Compound Document
NB. Excel 8.0 (Excel 97) BIFF8 Workbook Compound Document
NB. Excel 9.0 (Excel 2000) BIFF8 Workbook Compound Document
NB. Excel 10.0 (Excel XP) BIFF8X Workbook Compound Document
NB. Excel 11.0 (Excel 2003) BIFF8X Workbook Compound Document
NB. Index     Format String
NB. ---------------------------------------------------------
NB. 0 General General
NB. 1 Decimal 0
NB. 2 Decimal 0.00
NB. 3 Decimal #,##0
NB. 4 Decimal #,##0.00
NB. 5 Currency "$"#,##0_); ("$"#,##0)
NB. 6 Currency "$"#,##0_);[Red] ("$"#,##0)
NB. 7 Currency "$"#,##0.00_); ("$"#,##0.00)
NB. 8 Currency "$"#,##0.00_);[Red] ("$"#,##0.00)
NB. 9 Percent 0%
NB. 10 Percent 0.00%
NB. 11 Scientific 0.00E+00
NB. 12 Fraction # ?/?
NB. 13 Fraction # ??/??
NB. 14 Date M/D/YY
NB. 15 Date D-MMM-YY
NB. 16 Date D-MMM
NB. 17 Date MMM-YY
NB. 18 Time h:mm AM/PM
NB. 19 Time h:mm:ss AM/PM
NB. 20 Time h:mm
NB. 21 Time h:mm:ss
NB. 22 Date/Time M/D/YY h:mm
NB. 37 Account. _(#,##0_);(#,##0)
NB. 38 Account. _(#,##0_);[Red](#,##0)
NB. 39 Account. _(#,##0.00_);(#,##0.00)
NB. 40 Account. _(#,##0.00_);[Red](#,##0.00)
NB. 41 Currency _ ("$"* #,##0_);_ ("$"* (#,##0);_ ("$"* "-"_);_(@_)
NB. 42 Currency _(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)
NB. 43 Currency _ ("$"* #,##0.00_);_ ("$"* (#,##0.00);_ ("$"* "-"??_);_(@_)
NB. 44 Currency _(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)
NB. 45 Time mm:ss
NB. 46 Time [h]:mm:ss
NB. 47 Time mm:ss.0
NB. 48 Scientific ##0.0E+0
NB. 49 Text @
format0n=: 164  NB. reserved by excel
colorset0n=: 8   NB. reserved by excel
colorborder=: 16b40
colorpattern=: 16b41
colorfont=: 16b7fff
NB. cell ref 'AA5' => 4 26
A1toRC=: 3 : 0
y=. y.
assert. y e. '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ'
(26 #. (0 _1)+ _2&{. ' ABCDEFGHIJKLMNOPQRSTUVWXYZ'&i. c),~ <: 1&". y -. c=. y -. '0123456789'
)

toHeader=: toWORD0
UString=: 3 : 0
y=. y.
toucode0 u: y
)

SString=: 3 : 0
y=. y.
(1&u: ::]) y
)

toString=: 3 : 0
y=. y.
if. 131072= 3!:0 y do.
  toucode0 y
else.
  y
end.
)

toUString8=: 3 : 0
y=. y.
if. 131072= 3!:0 y do.
  (a.{~#y), (1{a.), toucode0 y
else.
  (a.{~#y), (0{a.), y
end.
)

toUString16=: 3 : 0
y=. y.
if. 131072= 3!:0 y do.
  (toWORD0 #y), (1{a.), toucode0 y
else.
  (toWORD0 #y), (0{a.), y
end.
)

toUString=: 3 : 0
y=. y.
if. 131072= 3!:0 y do.
  (1{a.), toucode0 y
else.
  (0{a.), y
end.
)

toStream=: 4 : 0
y=. y. [ x=. x.
x fappend~ y
)

sulen8=: 3 : 0
y=. y.
if. 131072= 3!:0 y do.
  2+2*#y
else.
  2+#y
end.
)

sulen16=: 3 : 0
y=. y.
if. 131072= 3!:0 y do.
  3+2*#y
else.
  3+#y
end.
)

NB. enum constants
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

NB. ---------------------------------------------------------

coinstance=: {. @ (18!:2 & boxxopen)

NB. utility for writing continue record
NB. biffappend
NB.
NB. General storage function
NB.

biffappend=: 3 : 0
y=. y.

0 biffappend y
:
y=. y. [ x=. x.

if. RECORDLEN < #y do. x add_continue y else. y end.
)

NB. add_continue
NB.
NB. Excel limits the size of BIFF records. In Excel 5 the limit is 2084 bytes. In
NB. Excel 97 the limit is 8228 bytes. Records that are longer than these limits
NB. must be split up into CONTINUE blocks.
NB.
NB. This function take a long BIFF record and inserts CONTINUE records as
NB. necessary.
NB.
NB. Some records have their own specialised Continue blocks so there is also an
NB. option to bypass this function.
NB.
add_continue=: 4 : 0
y=. y. [ x=. x.

data=. y
recordtype=. 16b003c NB. Record identifier

NB. Skip this if another method handles the continue blocks.
if. x do. data return. end.

NB. The first 2080/8224 bytes remain intact. However, we have to change
NB. the length field of the record.
data=. RECORDLEN}.data [ tmp=. RECORDLEN{.data
tmp=. (toWORD0 RECORDLEN-4) 2 3}tmp

NB. Strip out chunks of 2080/8224 bytes +4 for the header.
while. RECORDLEN < #data do.
  tmp=. tmp, (RECORDLEN{.data),~ toHeader recordtype, RECORDLEN
  data=. RECORDLEN}.data
end.

NB. Mop up the last of the data
tmp=. tmp, data,~ toHeader recordtype, #data
)

NB. extend crc32 to 16 bytes to simulate md4
biffcrc32=: 3 : 0
y=. y.

2&(3!:4) 4# (128!:3) y
)

biffmd4=: 3 : 0
y=. y.

md5bin_crypt_ y
)

NB. add_mso_generic
NB.
NB. Create a mso structure that is part of an Escher drawing object. These are
NB. are used for images, comments and filters. This generic method is used by
NB. other methods to create specific mso records.
NB.
NB. Returns the packed record.
NB.
add_mso_generic=: 4 : 0
y=. y. [ x=. x.

data=. x
'type version instance length'=. y

NB. The header contains version and instance info packed into 2 bytes.
header=. version (23 b.) 4 (33 b.) instance

record=. toWORD0 header, type
record=. record, toDWORD0 length
record, data
)

NB. ---------------------------------------------------------

NB. biff record
NB. a formula which was array-entered into a range of cells
biff_array=: 3 : 0
y=. y.
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

NB. Should Excel make a backup of this XLS sheet?
biff_backup=: 3 : 0
y=. y.
recordtype=. 16b0040
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)

NB. a cell with no formula or value
biff_blank=: 4 : 0
y=. y. [ x=. x.
recordtype=. 16b0201
y=. >y
assert. 2=#y
z=. ''
z=. z, toWORD0 y
z=. z, toWORD0 x
z=. z,~ toHeader recordtype, #z
)

NB. set the Beginning of File marker
NB.  version = Excel version   default is biff8 (97-2003)
NB.  docn = Excel document type   default is worksheet
biff_bof=: 3 : 0
y=. y.
'version docn'=. y
recordtype=. 16b809
z=. ''
z=. z, toWORD0 version
z=. z, toWORD0 docn
z=. z, toWORD0 8111   NB. built id
z=. z, toWORD0 1997   NB. built year
z=. z, toDWORD0 16b40c1
z=. z, toDWORD0 16b106
z=. z,~ toHeader recordtype, #z
)

biff_bookbool=: 3 : 0
y=. y.
recordtype=. 16b00da
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)

NB. a cell with a constant boolean OR error
NB. boolerrvalue = boolean or error value
NB. boolORerr = specifies a boolean or error   1 =. error, 0 =. boolean
biff_boolerr=: 4 : 0
y=. y. [ x=. x.
'rowcol boolerrvalue boolORerr'=. y
recordtype=. 16b0205
z=. ''
z=. z, toWORD0 rowcol
z=. z, toWORD0 x
z=. z, toBYTE boolerrvalue
z=. z, toBYTE boolORerr
z=. z,~ toHeader recordtype, #z
)

NB. set the bottom margin, used when printing, in inches
NB. 8-byte IEEE floating point format
biff_bottommargin=: 3 : 0
y=. y.
recordtype=. 16b0029
z=. ''
z=. z, toDouble0 y
z=. z,~ toHeader recordtype, #z
)

biff_boundsheet=: 3 : 0
y=. y.
'offset visible type sheetname'=. y
recordtype=. 16b85
z=. ''
z=. z, toDWORD0 offset
z=. z, toBYTE visible
z=. z, toBYTE type
z=. z, toUString8 sheetname
z=. z,~ toHeader recordtype, #z
)

NB. set the iteration count as set in Options->Calculation
biff_calccount=: 3 : 0
y=. y.
recordtype=. 16b000c
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)

NB. set the calculation mode
biff_calcmode=: 3 : 0
y=. y.
recordtype=. 13
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)

biff_codepage=: 3 : 0
y=. y.
recordtype=. 16b0042
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)

biff_colinfo=: 4 : 0
y=. y. [ x=. x.
'col1 col2 width hide level collapse'=. y
recordtype=. 16b007d
z=. ''
z=. z, toWORD0 col1, col2, width, x
z=. z, toWORD0 hide bitor 8 bitshl level bitor 4 bitshl collapse
z=. z, toWORD0 0
z=. z,~ toHeader recordtype, #z
)

NB. set the default cell attributes for those cells which aren't
NB. defined
biff_columndefault=: 3 : 0
y=. y.
recordtype=. 16b0020
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)

NB. a continuation of a formula too long to fit into one cell
biff_continue=: 3 : 0
y=. y.
recordtype=. 16b003c
z=. ''
z=. z, toString y
z=. z,~ toHeader recordtype, #z
)

biff_country=: 3 : 0
y=. y.
recordtype=. 16b008c
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)

biff_crn=: 3 : 0
y=. y.
recordtype=. 16b005a
z=. ''
z=. z, SString y
z=. z,~ toHeader recordtype, #z
)

NB. Specify the date system used in the XLS worksheet
biff_date1904=: 3 : 0
y=. y.
recordtype=. 16b0022
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)

NB. width is in character
biff_defaultcolwidth=: 3 : 0
y=. y.
recordtype=. 16b0055
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)

NB. the row height of all undefined rows
NB. does not affect explicitly defined rows
NB. height is in increments of 1/20th of a point
biff_defaultrowheight=: 3 : 0
y=. y.
'notmatch hidden spaceabove spacebelow height'=. y
recordtype=. 16b0225
z=. ''
z=. z, toWORD0 bitor (0~:notmatch) bitor 1 bitshl (0~:hidden) bitor 1 bitshl (0~:spaceabove) bitor 1 bitshl (0~:spacebelow)
z=. z, toWORD0 height
z=. z,~ toHeader recordtype, #z
)

NB. set the maximum change for an interative record
biff_delta=: 3 : 0
y=. y.
recordtype=. 16b0010
z=. ''
z=. z, toDouble0 y
z=. z,~ toHeader recordtype, #z
)

NB. the minimum and maximum bounds of the worksheet
NB. the last row and column are numbered one higher than
NB. the last occupied row/column
biff_dimensions=: 3 : 0
y=. y.
recordtype=. 16b0200
z=. ''
z=. z, toDWORD0 0 1+0 1{y
z=. z, toWORD0 0 1+2 3{y
z=. z, toWORD0 0
z=. z,~ toHeader recordtype, #z
)

NB. set the EOF record
biff_eof=: 3 : 0
y=. y.
recordtype=. 16b000a
z=. ''
z=. z,~ toHeader recordtype, #z
)

biff_externname=: 4 : 0
y=. y. [ x=. x.
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
y=. y.
recordtype=. 16b0017
z=. ''
z=. z, toWORD0 #y
for_yi. y do.
  z=. z, toWORD0 yi
end.
z=. z,~ toHeader recordtype, #z
)

NB. describes an entry in Excels font table
NB.  height of the font in 1/20 of a point increments
biff_font=: 3 : 0
y=. y.
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

NB. sepcify the footer for the worksheet
biff_footer=: 3 : 0
y=. y.
recordtype=. 16b0015
z=. ''
if. #y do.
  z=. z, toUString16 y
end.
z=. z,~ toHeader recordtype, #z
)

NB. describes a cell format
biff_format=: 3 : 0
y=. y.
'num str'=. y
recordtype=. 16b041e
z=. ''
z=. z, toWORD0 num
z=. z, toUString16 str
z=. z,~ toHeader recordtype, #z
)

NB. a cell with a formula
NB.  value = current value of the formula
NB.  recalc = should the cell be recalculated on XLS load?
NB.  exprlength = length of parsed expression
NB.  pasedexpr = parsed expression
biff_formula=: 4 : 0
y=. y. [ x=. x.
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
y=. y.
recordtype=. 16b0082
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)

biff_guts=: 3 : 0
y=. y.
recordtype=. 16b0080
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)

biff_hcenter=: 3 : 0
y=. y.
recordtype=. 16b008d
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)

NB. specify the header for the worksheet
biff_header=: 3 : 0
y=. y.
recordtype=. 16b0014
z=. ''
if. #y do.
  z=. z, toUString16 y
end.
z=. z,~ toHeader recordtype, #z
)

biff_hideobj=: 3 : 0
y=. y.
recordtype=. 16b008d
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)

biff_hlink=: 4 : 0
y=. y. [ x=. x.
recordtype=. 16b01b8
z=. ''
linktype=. x
if. 'url'-:linktype do. 'rowcols link textmark target description'=. y
elseif. 'local'-:linktype do. 'rowcols link elink uplevel textmark target description'=. y
elseif. 'unc'-:linktype do. 'rowcols link textmark target description'=. y
elseif. 'worksheet'-:linktype do. 'rowcols textmark target description'=. y
elseif. do. 'unhandled exception' 13!:8 (3)
end.
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
z=. z, toWORD0 rowcols  NB. rowcols is row1 row2 col1 col2
z=. z, SString 16bd0 a.{~16bc9 16bea 16b79 16bf9 16bba 16bce 16b11 16b8c 16b82 16b00 16baa 16b00 16b4b 16ba9 16b0b
z=. z, toDWORD0 2
z=. z, toDWORD0 flag
if. #description do.
  z=. z, toDWORD0 1+#description
  z=. z, UString description+{.a.
end.
if. #target do.
  z=. z, toDWORD0 1+#target
  z=. z, UString target+{.a.
end.
if. 'url'-:linktype do.
  z=. z, SString a.{~16be0 16bc9 16bea 16b79 16bf9 16bba 16bce 16b11 16b8c 16b82 16b00 16baa 16b00 16b4b 16ba9 16b0b
  z=. z, toDWORD0 2*1+#link
  z=. z, UString link+{.a.
elseif. 'local'-:linktype do.
  z=. z, SString a.{~16b03 16b03 16b00 16b00 16b00 16b00 16b00 16b00 16bc0 16b00 16b00 16b00 16b00 16b00 16b00 16b46
  z=. z, toWORD0 uplevel
  z=. z, toDWORD0 1+#link
  z=. z, SString link+{.a.
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
  z=. z, UString link+{.a.
elseif. 'worksheet'-:linktype do.
  ''  NB. only textmark present
end.
if. #textmark do.
  z=. z, toDWORD0 1+#textmark
  z=. z, UString textmark+{.a.
end.
z=. z,~ toHeader recordtype, #z
)

NB. a list of explicit row page breaks
biff_horizontalpagebreaks=: 3 : 0
y=. y.
recordtype=. 16b001b
z=. ''
z=. z, toWORD0 #y
for_iy. y do. z=. z, toWORD0 iy end.
z=. z,~ toHeader recordtype, #z
)

biff_index=: 3 : 0
y=. y.
recordtype=. 16b020b
z=. ''
z=. z, toDWORD0 0
z=. z, toDWORD0 y
z=. z,~ toHeader recordtype, #z
)

NB. a cell containing a 14-bit signed integer, biff8 use RK value to store integer
NB. negative numbers and those outside this range are to
NB. be held in a NUMBER cell
biff_integer=: 4 : 0
y=. y. [ x=. x.
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

NB. set the state of iteration
biff_iteration=: 3 : 0
y=. y.
recordtype=. 16b0011
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)

NB. cell with a constant string of length
biff_label=: 4 : 0
y=. y. [ x=. x.
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
y=. y.
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

NB. set the left margin, used when printing, in inches
NB. 8-byte IEEE floating point format
biff_leftmargin=: 3 : 0
y=. y.
recordtype=. 16b0026
z=. ''
z=. z, toDouble0 y
z=. z,~ toHeader recordtype, #z
)

biff_mergedcells=: 3 : 0
y=. y.
if. (0:=#) y do. '' return. end.
recordtype=. 16b00e5
z=. ''
z=. z, toWORD0 #y
for_iy. y do. z=. z, toWORD0 iy end.
z=. z,~ toHeader recordtype, #z
)

biff_name=: 3 : 0
y=. y.
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

NB. a cell containing a constant floating point number
NB. IEEE 8-byte floating point format
biff_number=: 4 : 0
y=. y. [ x=. x.
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

NB. workbook/worksheet object protection
biff_objectprotect=: 3 : 0
y=. y.
recordtype=. 16b0063
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)

biff_palette=: 3 : 0
y=. y.
recordtype=. 16b0092
z=. ''
z=. z, toWORD0 #y
z=. z, toDWORD0 y
z=. z,~ toHeader recordtype, #z
)

NB. number and position of panes in a window
biff_pane=: 3 : 0
y=. y.
recordtype=. 16b0041
z=. ''
'split vis activepane'=. y
z=. z, toWORD0 split   NB. xsplit ysplit
z=. z, toWORD0 vis     NB. topvis leftvis
for_pane. activepane do. z=. z, toWORD0 pane end.
z=. z,~ toHeader recordtype, #z
)

NB. worksheet/workbook protection (.5.18). It stores a 16-bit hash value,
biff_password=: 3 : 0
y=. y.
recordtype=. 16b0013
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)

NB. set the precision as set in Option->Calculation
biff_precision=: 3 : 0
y=. y.
recordtype=. 16b000e
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)

NB. Print Gridlines when printing?
biff_printgridlines=: 3 : 0
y=. y.
recordtype=. 16b002b
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)

NB. Shall Excel print the row 'n column headers
biff_printheaders=: 3 : 0
y=. y.
recordtype=. 16b002a
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)

NB. workbook/worksheet cell content protection
biff_protect=: 3 : 0
y=. y.
recordtype=. 16b0012
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)

NB. set the cell reference mode as in Options->Desktop
NB. sets the cell reference mode to
NB. <letter><number>   like A1 or C3 -- you sank my battleship
NB.    OR
NB. R<number>C<number>   y'know R1C1 = row 1 col 1
biff_refmode=: 3 : 0
y=. y.
recordtype=. 16b000f
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)

NB. set the right margin, used when printing, in inches
NB. 8-byte IEEE floating point format
biff_rightmargin=: 3 : 0
y=. y.
recordtype=. 16b0027
z=. ''
z=. z, toDouble0 y
z=. z,~ toHeader recordtype, #z
)

NB. a row descriptor needs the follwing ingredients
biff_row=: 4 : 0
y=. y. [ x=. x.
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

NB. workbook/worksheet scenarios protection
biff_scenprotect=: 3 : 0
y=. y.
recordtype=. 16b00dd
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)

biff_scl=: 3 : 0
y=. y.
recordtype=. 16b00a0
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)

NB. sets the cells which are selected in a pane
NB. not implemented yet
biff_selection=: 3 : 0
y=. y.
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
y=. y.
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
y=. y.
recordtype=. 16b0099
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)

NB. the string value of a formula
NB. the value of all formulas outside of this record are held in
NB. Excels formula format
biff_string=: 3 : 0
y=. y.
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
y=. y. [ x=. x.
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
y=. y. [ x=. x.
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

NB. a one-input row or column table created with the Data Table command
biff_table=: 3 : 0
y=. y.
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

NB. set the top margin, used when printing, in inches
NB. 8-byte IEEE floating point format
biff_topmargin=: 3 : 0
y=. y.
recordtype=. 16b0028
z=. ''
z=. z, toDouble0 y
z=. z,~ toHeader recordtype, #z
)

biff_vcenter=: 3 : 0
y=. y.
recordtype=. 16b0084
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)

NB. a list of explicit column page breaks
biff_verticalpagebreaks=: 3 : 0
y=. y.
recordtype=. 16b001a
z=. ''
z=. z, toWORD0 #y
for_iy. y do. z=. z, toWORD0 iy end.
z=. z,~ toHeader recordtype, #z
)

NB. basic Excel window information
biff_window1=: 3 : 0
y=. y.
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

NB. advanced window information
biff_window2=: 3 : 0
y=. y.
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

NB. workbook/worksheet window configuration protection
biff_windowprotect=: 3 : 0
y=. y.
recordtype=. 16b0019
z=. ''
z=. z, toWORD0 0~:y
z=. z,~ toHeader recordtype, #z
)

biff_wsbool=: 3 : 0
y=. y.
recordtype=. 16b0081
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)

biff_xct=: 3 : 0
y=. y.
recordtype=. 16b0059
z=. ''
z=. z, toWORD0 y
z=. z,~ toHeader recordtype, #z
)

biff_xf=: 3 : 0
y=. y.
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

coxclass 'biffrefname'
coxtend 'biff'
create=: 3 : 0
y=. y.
'hidden function command macro complex builtin functiongroup binaryname keybd sheetidx'=: 0
'name formula menu description helptopic statusbar'=: 6#a:
)

destroy=: codestroy
writestream=: 3 : 0
y=. y.
z=. biff_name hidden ; function ; command ; macro ; complex ; builtin ; functiongroup ; binaryname ; keybd ; name ; formula ; sheetidx ; menu ; description ; helptopic ; statusbar
)

coxclass 'biffsupbook'
coxtend 'biff'
newextname=: 3 : 0
y=. y.
extname=: extname, <y
)

create=: 3 : 0
y=. y.
'type sheetn source sheetname'=: 4{.y
if. -. (<type) e. 'external' ; 'internal' ; 'addin' ; 'ole' ; 'dde' do.
  'unhandled exception' 13!:8 (3)
end.
extname=: ''
crn=: ''  NB. not yet implemented
)

destroy=: codestroy
writestream=: 3 : 0
y=. y.
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

coxclass 'biffxf'
coxtend 'biff'
NB. merely create biffxf will not result in a new biff xf record in excel file
NB. each biffxf object must be getxfidx
NB. getcolor=: 3 : 0
NB. y=. y.
NB. if. y e. colorborder, colorpattern, colorfont do. y return. end.
NB. if. (#colorset__COCREATOR)= n=. colorset__COCREATOR i. y do.
NB.   colorset__COCREATOR=: colorset__COCREATOR, y
NB. end.
NB. n
NB. )
getcolor=: ]
getfont=: 3 : 0
y=. y.
'fontheight fontitalic fontstrike fontcolor fontweight fontscript fontunderline fontfamily fontcharset fontname'=. y
y=. fontheight ; fontitalic ; fontstrike ; (getcolor fontcolor) ; fontweight ; fontscript ; fontunderline ; fontfamily ; fontcharset ; dltb fontname
if. (#fontset__COCREATOR)= n=. fontset__COCREATOR i. y do.
  fontset__COCREATOR=: fontset__COCREATOR, y
end.
n
)

getformat=: 3 : 0
y=. y.
if. (#formatset__COCREATOR)= n=. formatset__COCREATOR i. <y=. dltb y do.
  formatset__COCREATOR=: formatset__COCREATOR, <y
end.
n
)

NB. return xf in xfset row for an biffxf object
NB. usage: getxfrow__xfo ''
NB. font format typeprotparent align rotate indent used border linecolor color
getxfrow=: 3 : 0
y=. y.
font=. getfont fontheight ; fontitalic ; fontstrike ; fontcolor ; fontweight ; fontscript ; fontunderline ; fontfamily ; fontcharset ; fontname
formatn=. getformat format
NB. see 6.70, otherwise excel cannot use cell format edit
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

NB. set biffxf object to an xfset row
NB. font format typeprotparent align rotate indent used border linecolor color
copyxfrow=: 3 : 0
y=. y.
'font formatn typeprotparent align rotate indentshrink used border linecolor color'=. y
'fontheight fontitalic fontstrike fontcolor fontweight fontscript fontunderline fontfamily fontcharset fontname'=: font{fontset__COCREATOR
NB. if. -. fontcolor e. colorborder, colorpattern, colorfont do.
NB.   fontcolor=: fontcolor{.colorset__COCREATOR
NB. end.
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

NB. copy content from an another biffxf object
copyxfobj=: 3 : 0
y=. y.
l=. y
nm=. nl__l 0
for_nmi. nm do. (>nmi)=: ((>nmi), '__l')~ end.
)

create=: 3 : 0
y=. y.
NB. read section 5.113 XF Extended Format and 5.43 FONT
'fontheight fontitalic fontstrike fontcolor fontweight fontscript fontunderline fontfamily fontcharset fontname'=: {.fontset_COCREATOR
format=: 'General'
lock=: 0
hideformula=: 1
type=: 0  NB. 0: cell 1:style
parentxf=: 0
horzalign=: 0  NB. 0 general  1 left  2 center  3 right
textwrap=: 0
vertalign=: 2  NB. 0 top  1 centre  2 bottom
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
xfindex=: _1   NB. undefined
if. ''-.@-:y do.
  if. (3!:0 y) e. 1 4 do.
    copyxfrow y
  elseif. (3!:0 y) e. 32 do.
    copyxfobj y
  end.
end.
)

destroy=: codestroy

coxclass 'biffsheet'
coxtend 'biff'

NB. This verb use IMDATA record which has been obsoleted, only MS Excel can display it correct.
NB. other Excel compatible program (open office calc, gnumeric can not display it.
NB.   Insert a 24bit bitmap image in a worksheet.
NB. x
NB.   row     The row we are going to insert the bitmap into
NB.   col     The column we are going to insert the bitmap into
NB.   x       The horizontal position (offset) of the image inside the cell.
NB.   y       The vertical position (offset) of the image inside the cell.
NB.   scale_x The horizontal scale
NB.   scale_y The vertical scale
NB. y
NB.   bitmap  The bitmap filename
insertpicture=: 4 : 0
y=. y. [ x=. x.
img=. y
'rowcol xyoffset scalexy'=. x
if. 2 131072 e.~ 3!:0 rowcol do. rowcol=. A1toRC rowcol end.
'row col'=. rowcol
'x1 y1'=. xyoffset
'scalex scaley'=. scalexy
z=. processbitmap img
if. _1&= 0{::z do. z return. end.
'width height size data'=. }.z
NB. Scale the frame of the image.
width=. width * scalex
height=. height * scaley
NB. Calculate the vertices of the image and write the OBJ record
positionImage row, col, x1, y1, width, height
NB. Write the IMDATA record to store the bitmap data
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

NB.
NB.   Calculate the vertices that define the position of the image as required by
NB.   the OBJ record.
NB.
NB.           +------------+------------+
NB.           |     A      |      B     |
NB.     +-----+------------+------------+
NB.     |     |(x1, y1)     |            |
NB.     |  1  |(A1)._______|______      |
NB.     |     |    |              |     |
NB.     |     |    |              |     |
NB.     +-----+----|    BITMAP    |-----+
NB.     |     |    |              |     |
NB.     |  2  |    |______________.     |
NB.     |     |            |        (B2)|
NB.     |     |            |     (x2, y2)|
NB.     +---- +------------+------------+
NB.
NB.   Example of a bitmap that covers some of the area from cell A1 to cell B2.
NB.
NB.   Based on the width and height of the bitmap we need to calculate 8 vars:
NB.       $col_start, $row_start, $col_end, $row_end, $x1, $y1, $x2, $y2.
NB.   The width and height of the cells are also variable and have to be taken into
NB.   account.
NB.   The values of $col_start and $row_start are passed in from the calling
NB.   function. The values of $col_end and $row_end are calculated by subtracting
NB.   the width and height of the bitmap from the width and height of the
NB.   underlying cells.
NB.   The vertices are expressed as a percentage of the underlying cell width as
NB.   follows (rhs values are in pixels):
NB.
NB.         x1 = X / W *1024
NB.         y1 = Y / H *256
NB.         x2 = (X-1) / W *1024
NB.         y2 = (Y-1) / H *256
NB.
NB.         Where:  X is distance from the left side of the underlying cell
NB.                 Y is distance from the top of the underlying cell
NB.                 W is the width of the cell
NB.                 H is the height of the cell
NB.
NB.   the SDK incorrectly states that the height should be expressed as a percentage of 1024.
NB.
NB.   notice relative ordering of row col x1 y1
NB.   row_start Row containing top left corner of object
NB.   col_start Col containing upper left corner of object
NB.   x1        Distance to left side of object
NB.   y1        Distance to top of object
NB.   width     Width of image frame
NB.   height    Height of image frame
NB.
position_object=: 3 : 0
y=. y.
'rowstart colstart x1 y1 width height'=. y
NB. Initialise end cell to the same as the start cell
colend=. colstart  NB. Col containing lower right corner of object
rowend=. rowstart  NB. Row containing bottom right corner of object
NB. Zero the specified offset if greater than the cell dimensions
if. x1 >: sizeCol colstart do. x1=. 0 end.
if. y1 >: sizeRow rowstart do. y1=. 0 end.
width=. width + x1 -1
height=. height + y1 -1
NB. Subtract the underlying cell widths to find the end cell of the image
while. width >: sizeCol colend do.
  width=. width - sizeCol colend
  colend=. >:colend
end.
NB. Subtract the underlying cell heights to find the end cell of the image
while. height >: sizeRow rowend do.
  height=. height - sizeRow rowend
  rowend=. >:rowend
end.
NB. Bitmap isn't allowed to start or finish in a hidden cell, i.e. a cell
NB. with zero height or width.
NB.
if. 0= sizeCol colstart do. return. end.
if. 0= sizeCol colend do. return. end.
if. 0= sizeRow rowstart do. return. end.
if. 0= sizeRow rowend do. return. end.
NB. Convert the pixel values to the percentage value expected by Excel
x1=. <. 1024 * x1 % sizeCol colstart
y1=. <. 256 * y1 % sizeRow rowstart
x2=. <. 1024 * width % sizeCol colend NB. Distance to right side of object
y2=. <. 256 * height % sizeRow rowend NB. Distance to bottom of object
colstart , x1 , rowstart , y1 , colend , x2 , rowend , y2
)

positionImage=: storeobjpicture@:position_object

getcolsizes=: 3 : 0
y=. y.
if. (#colsizes)= i=. y i.~ {.("1) colsizes do.
  _1
else.
  {: i{colsizes
end.
)

getrowsizes=: 3 : 0
y=. y.
if. (#rowsizes)= i=. y i.~ {.("1) rowsizes do.
  _1
else.
  {: i{rowsizes
end.
)

NB.
NB.   Convert the width of a cell from user's units to pixels. By interpolation
NB.   the relationship is: y = 7x +5. If the width hasn't been set by the user we
NB.   use the default value. If the col is hidden we use a value of zero.
NB.
NB.   return The width in pixels
NB.
sizeCol=: 3 : 0
y=. y.
NB. Look up the cell value to see if it has been changed
if. _1~: getcolsizes y do.
  if. 0= getcolsizes y do.
    0 return.
  else.
NB. user unit is 0.38 + number of character
NB. y = 8x +5  seem correct on my 800X600 screen
    <. 5 + 8 * 0.38 + 256%~ getcolsizes y return.
  end.
else.
  <. 5 + 8 * 0.38 + 256%~ defaultcolwidth * 256 return.
end.
)

NB.
NB.   Convert the height of a cell from user's units to pixels. By interpolation
NB.   the relationship is: y = 4/3x. If the height hasn't been set by the user we
NB.   use the default value. If the row is hidden we use a value of zero. (Not
NB.   possible to hide row yet).
NB.
NB.   return The width in pixels
NB.
sizeRow=: 3 : 0
y=. y.
NB. Look up the cell value to see if it has been changed
if. _1~: getrowsizes y do.
  if. 0= getrowsizes y do.
    0 return.
  else.
NB. user unit is point (=1/72")
    <. (4%3) * 20%~ getrowsizes y return.
  end.
else.
  <. (4%3) * 20%~ defaultrowheight return.
end.
)

NB.
NB.   Store the OBJ record that precedes an IMDATA record. This could be generalise
NB.   to support other Excel objects.
NB.
NB.   colL Column containing upper left corner of object
NB.   dxL  Distance from left side of cell
NB.   rwT  Row containing top left corner of object
NB.   dyT  Distance from top of cell
NB.   colR Column containing lower right corner of object
NB.   dxR  Distance from right of cell
NB.   rwB  Row containing bottom right corner of object
NB.   dyB  Distance from bottom of cell
NB.
storeobjpicture=: 3 : 0
y=. y.
'colL dxL rwT dyT colR dxR rwB dyB'=. y
record=. 16b005d   NB. Record identifier
length=. 16b003c   NB. Bytes to follow
cObj=. 16b0001   NB. Count of objects in file (set to 1)
OT=. 16b0008   NB. Object type. 8 =. Picture
id=. 16b0001   NB. Object ID
grbit=. 16b0614   NB. Option flags
cbMacro=. 16b0000   NB. Length of FMLA structure
Reserved1=. 16b0000   NB. Reserved
Reserved2=. 16b0000   NB. Reserved
icvBack=. 16b09     NB. Background colour
icvFore=. 16b09     NB. Foreground colour
fls=. 16b00     NB. Fill pattern
fAuto=. 16b00     NB. Automatic fill
icv=. 16b08     NB. Line colour
lns=. 16bff     NB. Line style
lnw=. 16b01     NB. Line weight
fAutoB=. 16b00     NB. Automatic border
frs=. 16b0000   NB. Frame style
cf=. 16b0009   NB. Image format, 9 =. bitmap
Reserved3=. 16b0000   NB. Reserved
cbPictFmla=. 16b0000   NB. Length of FMLA structure
Reserved4=. 16b0000   NB. Reserved
grbit2=. 16b0001   NB. Option flags
Reserved5=. 16b0000   NB. Reserved
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

NB.
NB.   Convert a 24 bit bitmap into the modified internal format used by Windows.
NB.   This is described in BITMAPCOREHEADER and BITMAPCOREINFO structures in the
NB.   MSDN library.
NB.
NB.   return Array with data and properties of the bitmap
NB.
processbitmap=: 3 : 0
y=. y.
raiseError=. ''
NB. Open file.
if. 32=3!:0 y do.
  data=. ]`(''"_)@.(_1&-:)@:fread y
else.
  data=. y
end.
NB. Check that the file is big enough to be a bitmap.
if. ((#data) <: 16b36) do.
  raiseError=. 'size error'
  goto_error.
end.
NB. The first 2 bytes are used to identify the bitmap.
identity=. 2{.data
if. (identity -.@-: 'BM') do.
  raiseError=. 'signiture error'
  goto_error.
end.
NB. Remove bitmap data: ID.
data=. 2}.data
NB. Read and remove the bitmap size. This is more reliable than reading
NB. the data size at offset 16b22.
NB.
size=. fromDWORD0 4{.data
data=. 4}.data
size=. size - 16b36 NB. Subtract size of bitmap header.
size=. size + 16b0c NB. Add size of BIFF header.
NB. Remove bitmap data: reserved, offset, header length.
data=. 12}.data
NB. Read and remove the bitmap width and height. Verify the sizes.
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
NB. Read and remove the bitmap planes and bpp data. Verify them.
planesandbitcount=. fromWORD0 4{.data
data=. 4}.data
if. (24 ~: 1{planesandbitcount) do.  NB. Bitcount
  raiseError=. 'not 24 bit color'
  goto_error.
end.
if. (1 ~: 0{planesandbitcount) do.
  raiseError=. 'contain more than 1 plane'
  goto_error.
end.
NB. Read and remove the bitmap compression. Verify compression.
compression=. fromDWORD0 4{.data
data=. 4}.data
if. (compression ~: 0) do.
  raiseError=. 'compression not supported'
  goto_error.
end.
NB. Remove bitmap data: data size, hres, vres, colours, imp. colours.
data=. 20}.data
NB. Add the BITMAPCOREHEADER data
header=. toDWORD0 16b000c
header=. header, toWORD0 width, height, 16b01, 16b18
data=. header, data
0 ; width ; height ; size ; data
return.
label_error.
_1 ; raiseError
)

insertbackground=: 3 : 0
y=. y.
z=. processbitmap y
if. _1&= 0{::z do. z return. end.
'width height size data'=. }.z
NB. Write the BITMAP record to store the bitmap data
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
y=. y.
dimensions=: ((1 3{dimensions)>. y) 1 3}dimensions
)

initsheet=: 3 : 0
y=. y.
NB. read section 4.2.6 Record Order in BIFF8
sheetname=: y
NB. = calculation setting block
calccount=: 100
calcmode=: 1
refmode=: 1
delta=: 0.001000
iteration=: 0
NB. = calculation setting block
printheaders=: 0
printgridlines=: 0
defaultrowheightnotmatch=: 1
defaultrowheighthidden=: 0
defaultrowheightspaceabove=: 0
defaultrowheightspacebelow=: 0
defaultrowheight=: 315  NB. flag, 1/20 point
wsbool=: 0
NB. = page setting block
horizontalpagebreaks=: 0 3$''         NB. row after pagebreak, 0, _1
verticalpagebreaks=: 0 3$''           NB. col after pagebreak, 0, _1
NB. command used in header and footer
NB. && The "&" character itself
NB. &L Start of the left section
NB. &C Start of the centred section
NB. &R Start of the right section
NB. &P Current page number
NB. &N Page count
NB. &D Current date
NB. &T Current time
NB. &A Sheet name (BIFF5-BIFF8)
NB. &F File name without path
NB. &Z File path without file name (BIFF8X)
NB. &G Picture (BIFF8X)
NB. &U Underlining on/off
NB. &E Double underlining on/off (BIFF5-BIFF8)
NB. &S Strikeout on/off
NB. &X Superscript on/off (BIFF5-BIFF8)
NB. &Y Subscript on/off (BIFF5-BIFF8)
NB. &"<fontname>" Set new font <fontname>
NB. &"<fontname>, <fontstyle>"
NB.     Set new font with specified style <fontstyle>. The style <fontstyle> is in most cases
NB.     one of Regular, Bold, Italic, or Bold Italic. But this setting is dependent on the
NB.     used font, it may differ (localised style names, or Standard, Oblique,...). (BIFF5-
NB.     BIFF8)
NB. &<fontheight>
NB.     Set font height in points (<fontheight> is a decimal value). If this command is followed
NB.     by a plain number to be printed in the header, it will be separated from the font height
NB.     with a space character.
header=: ''  NB. '&A'
footer=: 'Page &P of &N'
hcenter=: 0
vcenter=: 0
leftmargin=: 0.5
rightmargin=: 0.5
topmargin=: 0.5
bottommargin=: 0.75
NB. setup
NB. read section 5.87 SETUP
'setuppapersize setupscaling setupstartpage setupfitwidth setupfitheight setuprowmajor setupportrait setupinvalid setupblackwhite setupdraft setupcellnote setuporientinvalid setupusestartpage setupnoteatend setupprinterror setupdpi setupvdpi setupheadermargin setupfootermargin setupnumcopy'=: 0
setuprowmajor=: 1
setupportrait=: 1
setupinvalid=: 1
setupheadermargin=: 0.75
setupfootermargin=: 0.75
backgroundbitmap=: ''
NB. = page setting block
protect=: 1
windowprotect=: 1
objectprotect=: 1
scenprotect=: 1
password=: ''
NB.
defaultcolwidth=: 8
colinfoset=: 0 7$''
dimensions=: 0 0 0 0
NB. cell table
window2=: 16b6b6 0 0 10
sclnum=: sclden=: 1
'xsplit ysplit topvis leftvis'=: 0
activepane=: ''
NB.  standardwidth=: 8*256            NB. appear buggy
mergedcell=: 0 4$''               NB. row1 row2 col1 col2
rowlabelrange=: collabelrange=: 0 4$''
condformatstream=: ''
selection=: 0 5$''
hlink=: ''
imdata=: ''
dvstream=: ''
colsizes=: 0 2$''
rowsizes=: 0 2$''
NB. mso stuff
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
)

writestream=: 3 : 0
y=. y.
z=. biff_bof 16b600, 16
NB. = calculation setting block
p1=. #z
z=. z, biff_index (0 1+2{.dimensions), 0
z=. z, biff_calccount calccount
z=. z, biff_calcmode calcmode
z=. z, biff_refmode refmode
z=. z, biff_delta delta
z=. z, biff_iteration iteration
NB. = calculation setting block
z=. z, biff_printheaders printheaders
z=. z, biff_printgridlines printgridlines
z=. z, biff_defaultrowheight defaultrowheightnotmatch, defaultrowheighthidden, defaultrowheightspaceabove, defaultrowheightspacebelow, defaultrowheight
z=. z, biff_wsbool wsbool
NB. = page setting block
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
NB. = page setting block
if. (#password) *. protect +. windowprotect +. objectprotect +. scenprotect do.
  if. protect do. z=. z, biff_protect protect end.
  if. windowprotect do. z=. z, biff_windowprotect windowprotect end.
  if. objectprotect do. z=. z, biff_objectprotect objectprotect end.
  if. scenprotect do. z=. z, biff_scenprotect scenprotect end.
  z=. z, biff_password passwordhash password
end.
NB. amend INDEX record, set Absolute stream position of the DEFCOLWIDTH record
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
NB. standardwidth broken
NB. z=. z, biff_standardwidth standardwidth
if. (2:~:$$) mergedcell do. mergedcell=: _4]\, mergedcell end.
z=. z, biff_mergedcells mergedcell
z=. z, biff_labelranges rowlabelrange ; collabelrange
z=. z, condformatstream
for_item. hlink do. z=. z, (0{::item) biff_hlink }.item end.
z=. z, dvstream
z=. z, biff_eof ''
)

NB. 16 bit hash value for worksheet password (5.18.4)
passwordhash=: 3 : 0
y=. y.
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
y=. y.
stream=: ''
initsheet y
)

destroy=: codestroy

coxclass 'biffbook'
coxtend 'biff'
getxfobj=: 3 : 0
y=. y.
z=. _1  NB. error
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
y=. y.
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
y=. y.
sstn=: >:sstn
if. (#sst) = b=. sst i. y do. sst=: sst, y end.
b
)

WriteSST=: 3 : 0
y=. y.
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
y=. y.
sstbuf=: sstbuf, toWORD0 y
bufn=: bufn-2
)

wrtw=: 3 : 0
y=. y.
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
y=. y.
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
y=. y.
sstbuf=: (toWORD0 RECORDLEN-bufn) (wrtp+2+i.2)}sstbuf
wrtp=: #sstbuf
sstbuf=: sstbuf, toWORD0 16b003c, 0
bufn=: RECORDLEN
)

initbook=: 3 : 0
y=. y.
NB. read section 4.2.6 Record Order in BIFF8
fileprotectionstream=: ''
workbookprotectionstream=: ''
codepage=: 1200
window1=: 672 192 11004 6636 16b38
backup=: 0
hideobj=: 0
date1904=: 0
precision=: 1
bookbool=: 1
NB. height italic strike color weight script underline family charset fontname
fontset=: 0 0$''
fontset=: fontset, fontsize ; 0 ; 0 ; colorfont ; 400 ; 0 ; 0 ; 0 ; 0 ; fontname     NB. 0
fontset=: fontset, fontsize ; 0 ; 0 ; colorfont ; 400 ; 0 ; 0 ; 0 ; 0 ; fontname     NB. 1
fontset=: fontset, fontsize ; 0 ; 0 ; colorfont ; 400 ; 0 ; 0 ; 0 ; 0 ; fontname     NB. 2
fontset=: fontset, fontsize ; 0 ; 0 ; colorfont ; 400 ; 0 ; 0 ; 0 ; 0 ; fontname     NB. 3
fontset=: fontset, fontsize ; 0 ; 0 ; colorfont ; 400 ; 0 ; 0 ; 0 ; 0 ; fontname     NB. 4 (missing)
fontset=: fontset, fontsize ; 0 ; 0 ; colorfont ; 400 ; 0 ; 0 ; 0 ; 0 ; fontname     NB. 5
formatset=: format0n#a:
NB. format 0-22
formatset=: ('General';'0';'0.00';'#,##0';'#,##0.00';'"$"#,##0_); ("$"#,##0)';'"$"#,##0_);[Red] ("$"#,##0)';'"$"#,##0.00_); ("$"#,##0.00)';'"$"#,##0.00_);[Red] ("$"#,##0.00)';'0%';'0.00%';'0.00E+00';'# ?/?';'# ??/??';'M/D/YY';'D-MMM-YY';'D-MMM';'MMM-YY';'h:mm AM/PM';'h:mm:ss AM/PM';'h:mm';'h:mm:ss';'M/D/YY h:mm') (i.23) } formatset
NB. format 37-49
formatset=: ('_(#,##0_);(#,##0)';'_(#,##0_);[Red](#,##0)';'_(#,##0.00_);(#,##0.00)';'_(#,##0.00_);[Red](#,##0.00)';'_ ("$"* #,##0_);_ ("$"* (#,##0);_ ("$"* "-"_);_(@_)';'_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)';'_ ("$"* #,##0.00_);_ ("$"* (#,##0.00);_ ("$"* "-"??_);_(@_)';'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';'mm:ss';'[h]:mm:ss';'mm:ss.0';'##0.0E+0';'@') (37+i.13)}formatset
NB. user-defined
formatset=: formatset, 'd/m/yyyy';'#,##0.000';'#,##0.0000';'#,##0.000000000'
NB. font format typeprot align rotate indent used border linecolor color
xfset=: 0 10$''
NB. style XF
xfset=: xfset, 0 0 16bfff5 16b20 0 0 0 0 0 16b20c0 NB. all valid
xfset=: xfset, 1 0 16bfff5 16b20 0 0 16bf4 0 0 16b20c0 NB. only font valid
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
NB. cell XF (15) will become default current xf (21)
xfset=: xfset, 0 0 1 16b20 0 0 0 0 0 16b20c0 NB. all valid
NB. Excel written XF, must include because they are referred by styleset
NB. (16)
xfset=: xfset, 1 16b2b 16bfff5 16b20 0 0 16bf8 0 0 16b20c0 NB. only format valid
xfset=: xfset, 1 16b29 16bfff5 16b20 0 0 16bf8 0 0 16b20c0
xfset=: xfset, 1 16b2c 16bfff5 16b20 0 0 16bf8 0 0 16b20c0
xfset=: xfset, 1 16b2a 16bfff5 16b20 0 0 16bf8 0 0 16b20c0
xfset=: xfset, 1 9 16bfff5 16b20 0 0 16bf8 0 0 16b20c0
NB. user-defined XF start here
NB. (21)
biffxfset=: ''
NB. xf builtin id level name
styleset=: 0 5$''
styleset=: styleset, 16b10 ; 1 ; 3 ; 16bff ; ''
styleset=: styleset, 16b11 ; 1 ; 6 ; 16bff ; ''
styleset=: styleset, 16b12 ; 1 ; 4 ; 16bff ; ''
styleset=: styleset, 16b13 ; 1 ; 7 ; 16bff ; ''
styleset=: styleset, 16b00 ; 1 ; 0 ; 16bff ; ''
styleset=: styleset, 16b14 ; 1 ; 5 ; 16bff ; ''
NB. predefine    black white red green blue yellow magenta cyan
colorset=: 0 16bffffff 16bff 16bff00 16bff0000 16bffff 16bff00ff 16bffff00
NB. Sets extra colour palette to the Excel 97+ default.
colorset=: colorset, RGB 16b00 16b00 16b00  NB. 8
colorset=: colorset, RGB 16bff 16bff 16bff  NB. 9
colorset=: colorset, RGB 16bff 16b00 16b00  NB. 10
colorset=: colorset, RGB 16b00 16bff 16b00  NB. 11
colorset=: colorset, RGB 16b00 16b00 16bff  NB. 12
colorset=: colorset, RGB 16bff 16bff 16b00  NB. 13
colorset=: colorset, RGB 16bff 16b00 16bff  NB. 14
colorset=: colorset, RGB 16b00 16bff 16bff  NB. 15
colorset=: colorset, RGB 16b80 16b00 16b00  NB. 16
colorset=: colorset, RGB 16b00 16b80 16b00  NB. 17
colorset=: colorset, RGB 16b00 16b00 16b80  NB. 18
colorset=: colorset, RGB 16b80 16b80 16b00  NB. 19
colorset=: colorset, RGB 16b80 16b00 16b80  NB. 20
colorset=: colorset, RGB 16b00 16b80 16b80  NB. 21
colorset=: colorset, RGB 16bc0 16bc0 16bc0  NB. 22
colorset=: colorset, RGB 16b80 16b80 16b80  NB. 23
colorset=: colorset, RGB 16b99 16b99 16bff  NB. 24
colorset=: colorset, RGB 16b99 16b33 16b66  NB. 25
colorset=: colorset, RGB 16bff 16bff 16bcc  NB. 26
colorset=: colorset, RGB 16bcc 16bff 16bff  NB. 27
colorset=: colorset, RGB 16b66 16b00 16b66  NB. 28
colorset=: colorset, RGB 16bff 16b80 16b80  NB. 29
colorset=: colorset, RGB 16b00 16b66 16bcc  NB. 30
colorset=: colorset, RGB 16bcc 16bcc 16bff  NB. 31
colorset=: colorset, RGB 16b00 16b00 16b80  NB. 32
colorset=: colorset, RGB 16bff 16b00 16bff  NB. 33
colorset=: colorset, RGB 16bff 16bff 16b00  NB. 34
colorset=: colorset, RGB 16b00 16bff 16bff  NB. 35
colorset=: colorset, RGB 16b80 16b00 16b80  NB. 36
colorset=: colorset, RGB 16b80 16b00 16b00  NB. 37
colorset=: colorset, RGB 16b00 16b80 16b80  NB. 38
colorset=: colorset, RGB 16b00 16b00 16bff  NB. 39
colorset=: colorset, RGB 16b00 16bcc 16bff  NB. 40
colorset=: colorset, RGB 16bcc 16bff 16bff  NB. 41
colorset=: colorset, RGB 16bcc 16bff 16bcc  NB. 42
colorset=: colorset, RGB 16bff 16bff 16b99  NB. 43
colorset=: colorset, RGB 16b99 16bcc 16bff  NB. 44
colorset=: colorset, RGB 16bff 16b99 16bcc  NB. 45
colorset=: colorset, RGB 16bcc 16b99 16bff  NB. 46
colorset=: colorset, RGB 16bff 16bcc 16b99  NB. 47
colorset=: colorset, RGB 16b33 16b66 16bff  NB. 48
colorset=: colorset, RGB 16b33 16bcc 16bcc  NB. 49
colorset=: colorset, RGB 16b99 16bcc 16b00  NB. 50
colorset=: colorset, RGB 16bff 16bcc 16b00  NB. 51
colorset=: colorset, RGB 16bff 16b99 16b00  NB. 52
colorset=: colorset, RGB 16bff 16b66 16b00  NB. 53
colorset=: colorset, RGB 16b66 16b66 16b99  NB. 54
colorset=: colorset, RGB 16b96 16b96 16b96  NB. 55
colorset=: colorset, RGB 16b00 16b33 16b66  NB. 56
colorset=: colorset, RGB 16b33 16b99 16b66  NB. 57
colorset=: colorset, RGB 16b00 16b33 16b00  NB. 58
colorset=: colorset, RGB 16b33 16b33 16b00  NB. 59
colorset=: colorset, RGB 16b99 16b33 16b00  NB. 60
colorset=: colorset, RGB 16b99 16b33 16b66  NB. 61
colorset=: colorset, RGB 16b33 16b33 16b99  NB. 62
colorset=: colorset, RGB 16b33 16b33 16b33  NB. 63
country=: 1 1
supbook=: ''
externsheet=: 0 3$''
extname=: ''
refname=: ''
NB. mso stuff
datasize=: 0
biffsize=: 0
mso_size=: 0
mso_clusters=: 0 5$''
images_size=: 0
images_data=: 0 0$''  NB. Store the data for MSODRAWINGGROUP.
)

addhlink=: 3 : 0
y=. y.
l=. sheeti{sheet
hlink__l=: hlink__l, y
)

writefileprotection=: 3 : 0
y=. y.
fileprotectionstream=: y
)

writeworkbookprotection=: 3 : 0
y=. y.
workbookprotectionstream=: y
)

writecelltable=: 3 : 0
y=. y.
l=. sheeti{sheet
stream__l=: stream__l, y
)

writecondformattable=: 3 : 0
y=. y.
l=. sheeti{sheet
condformatstream__l=: y
)

writedvtable=: 3 : 0
y=. y.
l=. sheeti{sheet
dvstream__l=: y
)

celladr=: 4 : 0
y=. y. [ x=. x.
'relrow relcol'=. x
'row col'=. y
z=. toWORD0 row
z=. z, toWORD0 (16bff bitand col) bitor 14 bitshl (0~:relcol) bitor 1 bitshl 0~:relrow
)

cellrangeadr=: 4 : 0
y=. y. [ x=. x.
'relrow1 relrow2 relcol1 relcol2'=. x
'row1 row2 col1 col2'=. y
z=. toWORD0 row1, row2
z=. z, toWORD0 (16bff bitand col1) bitor 14 bitshl (0~:relcol1) bitor 1 bitshl 0~:relrow1
z=. z, toWORD0 (16bff bitand col2) bitor 14 bitshl (0~:relcol2) bitor 1 bitshl 0~:relrow2
)

newsupbook=: 3 : 0
y=. y.
supbook=: supbook, y coxnew 'biffsupbook'
supbooki=: <:#supbook
)

newrefname=: 3 : 0
y=. y.
refname=: refname, '' coxnew 'biffrefname'
{:refname
)

newextname=: 3 : 0
y=. y.
l=. supbooki{supbook
newextname__l y
)

newcrn=: 3 : 0
y=. y.
l=. supbooki{supbook
newcrn__l y
)

create=: 3 : 0
y=. y.
if. ''-:y do.
  'fontname fontsize'=: 'Courier New' ; 220
  sheetname=. 'Sheet1'
else.
  'fontname fontsize'=: 2{.y
  sheetname=. 0{:: (2}.y) , <'Sheet1'
end.
sstn=: #sst=: ''
xfset=: sheet=: ''
initbook ''
addsheet sheetname  NB. create worksheet object
cxf=: addxfobj 15{xfset  NB. predefined cell style
)

destroy=: 3 : 0
y=. y.
for_l. sheet do. destroy__l '' end.
for_l. biffxfset do. destroy__l '' end.
for_l. supbook do. destroy__l '' end.
for_l. refname do. destroy__l '' end.
codestroy ''
)

NB. match color (red, green, blue) to internal color palette table
NB. return color index
rgbcolor=: 3 : 0
y=. y.
(i: <./) +/("1) | y -("1) RGBtuple colorset
)

NB. add new extended format
NB. return xf object
addxfobj=: 3 : 0
y=. y.
biffxfset=: biffxfset, z=. y coxnew 'biffxf'
z
)

NB. add new worksheet
NB. return worksheet index
NB. y sheet name or ''
addsheet=: 3 : 0
y=. y.
sheet=: sheet, ((y-:''){:: y ; 'Sheet', ": >:#sheet) coxnew 'biffsheet'
sheeti=: <:#sheet
)

NB. save to file
NB. y  filename ('' if return character data)
save=: 3 : 0
y=. y.
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
NB. excel peculiar, 4 is missing
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

NB. Calculate the offsets required by the BOUNDSHEET records
NB. calc_sheet_offsets ''

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
stream=. ('Workbook' ; '' ; '') coxnew 'oleppsfile'
append__stream olehead
append__stream olesheet
root=. (0 ; 0 ; <stream) coxnew 'oleppsroot'
rc=. save__root fn ; 0 ; ''
destroy__root ''
destroy__stream ''
rc
)

NB. set column width of current worksheet
NB. x xf
NB. y col1 col2 width hide level collapse
NB.    eg. 2 5 400  to set (col 2 3 4 5) 400 twip hight
addcolinfo=: 3 : 0
y=. y.
cxf addcolinfo y
:
y=. y. [ x=. x.
'col1 col2 width hide level collapse'=. 6{.y
l=. sheeti{sheet
colsizes__l=: ~. colsizes__l, (col1 + i.>:col2-col1) ,("0) (0~:hide){width, 0
colinfoset__l=: colinfoset__l, (getxfidx x), col1, col2, width, hide, level, collapse
''
)

NB. set row height of current worksheet
NB. x xf
NB. y rownumber firstcol_lastcol usedefaultheight rowheight heightnotmatch spaceabove spacebelow hidden explicitformat outlinelevel outlinegroup
NB.    eg. 3 0 256 0 12*256  to set (row 3) 12 character wide
addrowinfo=: 3 : 0
y=. y.
cxf addrowinfo y
:
y=. y. [ x=. x.
'rownumber firstcol lastcols usedefaultheight rowheight'=. 5{. y=. 12{.y
l=. sheeti{sheet
if. 0=usedefaultheight do.
  rowsizes__l=: ~. rowsizes__l, rownumber, rowheight
end.
stream__l=: stream__l, (getxfidx x) biff_row y
''
)

NB. write string to the current worksheet
NB. x xf
NB. y row col ; text         (where 3>$$text)
NB.    row col ; boxed text   (where 3>$$boxed text)
NB.                           (always box last argument to make 2=#y)
writestring=: 3 : 0
y=. y.
cxf writestring y
:
y=. y. [ x=. x.
assert. 2=#y
assert. 1 2 4 8 131072 e.~ (3!:0) 0{::y
assert. 2 32 131072 e.~ (3!:0) 1{::y
if. 2 131072 e.~ 3!:0 rc=. 0{::y do. y=. (<A1toRC rc) 0}y end.
l=. sheeti{sheet
xf=. getxfidx x
if. (0=#@, yn) +. 2 131072 e.~ 3!:0 yn=. 1{::y do.
  if. 2> $$yn do.
    adjdim__l 0{::y
    stream__l=: stream__l, xf biff_label y
  elseif. 2=$$yn do.
    'r c'=. 0{::y
    'nrow len'=. $yn
    adjdim__l 0{::y
    adjdim__l (nrow, 0) + 0{::y
    if. 0=len do.
      stream__l=: stream__l,, (toHeader 16b0201, 6) ,("1) (_2]\ toWORD0 r+i.nrow) ,("1) (_2]\ toWORD0 nrow#c) ,("1) (toWORD0 xf)
    else.
      yn=. <("1) yn
      sst=: sst, (~.yn) -. sst
      sstn=: sstn + #yn
      stream__l=: stream__l,, (toHeader 16b00fd, 10) ,("1) (_2]\ toWORD0 r+i.nrow) ,("1) (_2]\ toWORD0 nrow#c) ,("1) (toWORD0 xf) ,("1) (_4]\ toDWORD0 sst i. yn)
    end.
  elseif. do. 'unhandled exception' 13!:8 (3)
  end.
elseif. 32 e.~ 3!:0 yn do.
  if. (0:=#), yn do. '' return.  NB. ignore null
  elseif. 1=#@, yn do.  NB. singleton
    adjdim__l 0{::y
    stream__l=: stream__l, xf biff_label ({.y), <, >yn
  elseif. 3>$$yn do.
    if. 1=$$yn do. yn=. ,:yn end.
    'r c'=. 0{::y
    adjdim__l 0{::y
    adjdim__l (s=. $yn) + 0{::y
    if. #a=. I. > ''&-:&.> yn=. , yn do.
      yn=. ((#a)#<, ' ') a}yn   NB. biff8 cannot store empty string
    end.
    sst=: sst, (~.yn) -. sst
    sstn=: sstn + #yn
    stream__l=: stream__l,, (toHeader 16b00fd, 10) ,("1) (_2]\ toWORD0 ({:s)#r+i.{.s) ,("1) (_2]\ toWORD0, ({.s)#,:c+i.{:s) ,("1) (toWORD0 xf) ,("1) (_4]\ toDWORD0 sst i. yn)
  elseif. do. 'unhandled exception' 13!:8 (3)
  end.
elseif. do. 'unhandled exception' 13!:8 (3)
end.
''
)

NB. write integer to the current worksheet
NB. x xf
NB. y row col ; integer   (where 3>$$integer)
writeinteger=: 3 : 0
y=. y.
cxf writeinteger y
:
y=. y. [ x=. x.
assert. 2=#y
assert. 1 2 4 8 131072 e.~ (3!:0) 0{::y
assert. 1 4 8 e.~ (3!:0) 1{::y
if. 2 131072 e.~ 3!:0 rc=. 0{::y do. y=. (<A1toRC rc) 0}y end.
l=. sheeti{sheet
xf=. getxfidx x
NB. only 30 bit is used 536870911 = <:2^29
if. 536870911 < >./ |, 1{::y do. x writenumber y end.
if. (0:=#), yn=. 2b10 bitor 2 bitshl <. 1{::y do. '' return. end.  NB. ignore null
if. 1=#@, yn do.  NB. singleton
  adjdim__l 0{::y
  stream__l=: stream__l, xf biff_integer ({.y), < {., 1{::y
elseif. 3>$$yn do.
  if. 1=$$yn do. yn=. ,:yn end.
  'r c'=. 0{::y
  adjdim__l 0{::y
  adjdim__l (s=. $yn) + 0{::y
  stream__l=: stream__l,, (toHeader 16b027e, 10) ,("1) (_2]\ toWORD0 ({:s)#r+i.{.s) ,("1) (_2]\ toWORD0, ({.s)#,:c+i.{:s) ,("1) (toWORD0 xf) ,("1) (_4]\ toDWORD0, yn)
elseif. do. 'unhandled exception' 13!:8 (3)
end.
''
)

NB. write number to the current worksheet
NB. x xf
NB. y row col ; number   (where 3>$$number)
writenumber=: 3 : 0
y=. y.
cxf writenumber y
:
y=. y. [ x=. x.
assert. 2=#y
assert. 1 2 4 8 131072 e.~ (3!:0) 0{::y
assert. 1 4 8 e.~ (3!:0) 1{::y
if. 2 131072 e.~ 3!:0 rc=. 0{::y do. y=. (<A1toRC rc) 0}y end.
l=. sheeti{sheet
xf=. getxfidx x
if. (0:=#), yn=. 1{::y do. '' return. end.  NB. ignore null
if. 1=#@, yn do.  NB. singleton
  adjdim__l 0{::y
  stream__l=: stream__l, xf biff_number ({.y), < {., yn
elseif. 3>$$yn do.
  if. 1=$$yn do. yn=. ,:yn end.
  'r c'=. 0{::y
  adjdim__l 0{::y
  adjdim__l (s=. $yn) + 0{::y
  stream__l=: stream__l,, (toHeader 16b0203, 14) ,("1) (_2]\ toWORD0 ({:s)#r+i.{.s) ,("1) (_2]\ toWORD0, ({.s)#,:c+i.{:s) ,("1) (toWORD0 xf) ,("1) (_8]\ toDouble0, yn)
elseif. do. 'unhandled exception' 13!:8 (3)
end.
''
)

NB. write number to the current worksheet with format
NB. x xf
NB. y row col ; number ; format   (where 3>$$number)
writenumber2=: 3 : 0
y=. y.
cxf writenumber2 y
:
y=. y. [ x=. x.
assert. 3=#y
assert. 1 2 4 8 131072 e.~ (3!:0) 0{::y
assert. 1 4 8 e.~ (3!:0) 1{::y
assert. 2 131072 e.~ (3!:0) 2{::y
if. 2 131072 e.~ 3!:0 rc=. 0{::y do. y=. (<A1toRC rc) 0}y end.
l=. getxfobj x
t=. format__l
format__l=: _1{::y
if. 8= (3!:0) 1{::y do.
  l writenumber 2{.y
else.
  l writeinteger 2{.y
end.
format__l=: t
''
)

NB. write date to the current worksheet with format
NB. x xf
NB. y row col ; datenumber ; format   (where 3>$$datenumber)
writedate=: 3 : 0
y=. y.
cxf writedate y
:
y=. y. [ x=. x.
assert. 2 3 e.~ #y
assert. 1 2 4 8 131072 e.~ (3!:0) 0{::y
assert. 1 4 8 e.~ (3!:0) 1{::y
if. 2 131072 e.~ 3!:0 rc=. 0{::y do. y=. (<A1toRC rc) 0}y end.
l=. getxfobj x
t=. format__l
if. 2=#y do. y=. y, <shortdatefmt end. NB. default date format
assert. 2 131072 e.~ (3!:0) 2{::y
format__l=: 2{::y
l writenumber ({.y), <36522-~ 1{::y
format__l=: t
''
)

NB. insert bitmap to the current worksheet
NB. x (row, col) ; (offsetx(in point, 1/72"), offsety(in character)) ; scalex, scaley
NB. y boxed bitmap file name or charachar data of bitmap
NB. return if success   0 ; ''
NB.        if fail     _1 ; 'reason for failure'
insertpicture=: 4 : 0
y=. y. [ x=. x.
l=. sheeti{sheet
x insertpicture__l y
)

NB. insert bitmap as background for the current worksheet
NB. y boxed bitmap file name or charachar data of bitmap
NB. return if success   0 ; ''
NB.        if fail     _1 ; 'reson for failure'
insertbackground=: 3 : 0
y=. y.
l=. sheeti{sheet
insertbackground__l y
)

NB. builtin name
NB. 00H Consolidate_Area
NB. 01H Auto_Open
NB. 02H Auto_Close
NB. 03H Extract
NB. 04H Database
NB. 05H Criteria
NB. 06H Print_Area
NB. 07H Pint_Titles
NB. 08H Recorder
NB. 09H Data_Form
NB. 0AH Auto_Activate (BIFF5-BIFF8)
NB. 0BH Auto_Deactivate (BIFF5-BIFF8)
NB. 0CH Sheet_Title (BIFF5-BIFF8)
NB. 0DH _FilterDatabase (BIFF8)
NB. tmem 16b29
NB. tArea 16b3b
NB. tCellrangelist 16b10
NB. x supbook
NB.    _1 first call printarea, rowrepeat, colrepeat, rowcolrepeat
NB. y row1 row2 col1 col2
NB. return  supbook
printarea=: 4 : 0
y=. y. [ x=. x.
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

NB. x supbook
NB.    _1 first call printarea, rowrepeat, colrepeat, rowcolrepeat
NB. y sheetidx row1 row2
NB. return  supbook
rowrepeat=: 4 : 0
y=. y. [ x=. x.
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

NB. x supbook
NB.    _1 first call printarea, rowrepeat, colrepeat, rowcolrepeat
NB. y sheetidx col1 col2
NB. return  supbook
colrepeat=: 4 : 0
y=. y. [ x=. x.
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

NB. x supbook
NB.    _1 first call printarea, rowrepeat, colrepeat, rowcolrepeat
NB. y sheetidx row1 row2 col1 col2
NB. return  supbook
rowcolrepeat=: 4 : 0
y=. y. [ x=. x.
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

NB. ---------------------------------------------------------
NB. read numeric and string data from excel files
NB. written by bill lam
NB. ---------------------------------------------------------

coxclass 'biffread'
coxtend 'oleutlfcn'
NB. x 0 normal  1 debug
NB. y stream data
create=: 4 : 0
y=. y. [ x=. x.
debug=: x
stream=: y
biffver=: 0        NB. biff5/7 16b500   biff8 16b600
NB. biff2 16b200 biff3 16b300  biff4 16b400  (there is in fact no biffver stored)
streamtype=: 16    NB. worksheet 16
filepos=: 0
worksheets=: 0 2$''
bkrecords=: 0 3$'' [ bkbytes=: ''   NB. dump workbook for debug mode
records=: 0 3$'' [ bytes=: ''       NB. dump worksheet for debug mode
sst=: ''
bnsst=: nsst=: 0  NB. number of strings
sstchar=: 0
sstrtffe=: 0
insst=. 0
type=. _1
newptr=. 0
while. (type~:10)*.filepos<#stream do.
  'type ptr len'=. nextrecord ''
NB. dump records of workbook global to [bkbytes] with index held in [bkrecords]
  if. debug do.
    bkbytes=: bkbytes, (ptr+i.len){stream
    newptr=. newptr+len [ bkrecords=: bkrecords, type, newptr, len
  end.
  if. 16b9=type do.        NB. BOF record biff2
    biffver=: 16b200 break.
  elseif. 16b209=type do.  NB. BOF record biff3
    biffver=: 16b300 break.
  elseif. 16b409=type do.  NB. BOF record biff4
    biffver=: 16b400 break.
  elseif. 16b809=type do.  NB. BOF record
    if. 0=biffver do.  NB. only read the 1st BOF record in workbook
      'biffver streamtype'=: fromWORD0 4{.data=. (ptr+i.len){stream
    end.
  elseif. 16b0085=type do.  NB. boundsheet record
    if. 0={.fromBYTE 5{data=. (ptr+i.len){stream do.
      if. 16b500=biffver do.
        worksheets=: worksheets, (0 decodestring8 6}.data) ,< {.(fromDWORD0 4{.data)
      else.
        worksheets=: worksheets, (0 decodeustring8 6}.data) ,< {.(fromDWORD0 4{.data)
      end.
    end.
  elseif. 16b00fc=type do.  NB. sst record
    'bnsst nsst'=: fromDWORD0 8{.data=. (ptr+i.len){stream
    readsst 8}.data
    insst=. 1
  elseif. insst *. 16b003c=type do. NB. continue record after sst
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
NB. parse string from sst record
readsst=: 3 : 0
y=. y.
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

NB. return record type, pointer to data buffer, length of data buffer
nextrecord=: 3 : 0
y=. y.
'type len'=. fromWORD0 (filepos+i.4){stream
filepos=: +/data=. (4+filepos), len
type, data
)

NB. byte string used in biff2/3/4/5/7
NB. assume no continue record
decodestring8=: 4 : 0
y=. y. [ x=. x.
len=. {.fromBYTE (x+0){y
< len (([ <. #@]) {. ]) (x+1)}.y
)

decodestring16=: 4 : 0
y=. y. [ x=. x.
len=. {.fromWORD0 (x+i.2){y
< len (([ <. #@]) {. ]) (x+2)}.y
)

NB. assume no continue record
decodeustring8=: 4 : 0
y=. y. [ x=. x.
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
y=. y. [ x=. x.
NB. len number of (unicode) character in string
NB. z decoded string
NB. p point to rtf/fe block
NB. p1 byte length of decoded string segment
NB. p2 number of pending bytes to be read in next record
NB. lenrtf rtf size in byte (4 bytes per block)
NB. lenfe  fe size in byte
len=. {.fromWORD0 (x+i.2){y
uc=. 0~:1 bitand op=. {.fromBYTE (x+2){y  NB. uncompress unicode
fe=. 0~:4 bitand op                         NB. #far east phonetic (4 byte)
rtf=. 0~:8 bitand op                        NB. #rtf format run (2 byte)
lenrtffe=. 0
if. rtf do. lenrtffe=. 4* {.fromWORD0 ((x+3)+i.2){y end.
if. fe do. lenrtffe=. lenrtffe + {.fromDWORD0 ((x+3+(rtf*2))+i.4){y end.
NB. p point to position of fe/rtf block if read
NB. p2 #byte in character array not yet read
l=. (3+(fe*4)+(rtf*2)) + (len*uc{1 2) + lenrtffe  NB. expect length if no further continue record need
if. uc do.
  z=. fromucode0 z1=. (2*len) (([ <. #@]) {. ]) (x+3+(fe*4)+(rtf*2))}.y
  z2=. (2*len)-#z1  NB. byte length of remaining character array
else.
  z=. z1=. len (([ <. #@]) {. ]) (x+3+(fe*4)+(rtf*2))}.y
  z2=. len-#z1  NB. byte length of remaining character array
end.
if. (#y)<x+l do.
  z ; (#y), len, (x+l)-z2+#y
else.
  z ; l, len, 0
end.
)

NB. the first splitted string in continue record
decodeustring16a=: 4 : 0
y=. y. [ x=. x.
'x len lenrtffe'=. x
NB. there is no len for the first splitted string in continue record
NB. there is option byte the first splitted string in continue record
NB. but if length of remaining character array is zero, ie rtf/fe portion only
NB. then there will be no option byte
if. len do.
  uc=. 0~:1 bitand op=. {.fromBYTE (x+0){y
  l=. 1 + (len*uc{1 2) + lenrtffe  NB. expect length if no further continue record need
  if. uc do.
    z=. fromucode0 z1=. (2*len) (([ <. #@]) {. ]) (x+1)}.y
    z2=. (2*len)-#z1    NB. byte length of remaining character array
  else.
    z=. z1=. len (([ <. #@]) {. ]) (x+1)}.y
    z2=. len-#z1  NB. byte length of remaining character array
  end.
else.
  l=. lenrtffe  NB. expect length if no further continue record need
  z=. ''
  z2=. 0
end.
if. (#y)<x+l do.
  z ; (#y), len, (x+l)-z2+#y
else.
  z ; l, len, 0
end.
)

NB. x 1 convert all number to string
readsheet=: 0&$: : (4 : 0)
y=. y. [ x=. x.
if. 16b200 16b300 16b400 e.~ biffver do.  NB. biff2/3/4
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
records=: 0 3$'' [ bytes=: ''       NB. dump worksheet for debug mode
newptr=. 0
lookstr=. 0
while. filepos<#stream do.
  'type ptr len'=. nextrecord ''
NB. dump records of worksheet to [bytes] with index held in [records]
  if. debug do.
    bytes=: bytes, (ptr+i.len){stream
    newptr=. newptr+len [ records=: records, type, newptr, len
  end.
  if. 0=sheetfound do.  NB. biff2/3/4
    if. 16b0009 16b0209 16b0409 e.~ type do.  NB. biff2/3/4  BOF
      if. 16b10= fromWORD0 (2+ptr+i.2){stream do.  NB. worksheet stream
        if. y = scnt=. >:scnt do.         NB. test worksheet index wanted
          sheetfound=. 1
        end.
      end.
    end.
    continue.
  end.
  select. type
  case. 16b000a do. NB. EOF
    break.
  case. 16b027e do. NB. rk
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
      cellval=. cellval, ; null&,@ (":(!.maxpp))&.> getrk 4}.data   NB. ieee double have at most 16 sig. digits
    end.
  case. 16b0002 do. NB. unsigned 16-bit integer biff2
    if. 0=x do.
      rowcol4=. rowcol4, fromWORD0 4{.data=. (ptr+i.len){stream
      cellval4=. cellval4, {.fromDWORD0 (2#{.a.),~ (7+i.2){data
    else.
      rowcol=. rowcol, fromWORD0 4{.data=. (ptr+i.len){stream
      cellval=. cellval, null&,@ (":(!.maxpp)) {.fromDWORD0 (2#{.a.),~ (7+i.2){data
    end.
  case. 16b0003 do. NB. number biff2
    if. 0=x do.
      rowcol8=. rowcol8, fromWORD0 4{.data=. (ptr+i.len){stream
      cellval8=. cellval8, {.fromDouble0 (7+i.8){data
    else.
      rowcol=. rowcol, fromWORD0 4{.data=. (ptr+i.len){stream
      cellval=. cellval, null&,@ (":(!.maxpp)) {.fromDouble0 (7+i.8){data
    end.
  case. 16b0203 do. NB. number
    if. 0=x do.
      rowcol8=. rowcol8, fromWORD0 4{.data=. (ptr+i.len){stream
      cellval8=. cellval8, {.fromDouble0 (6+i.8){data
    else.
      rowcol=. rowcol, fromWORD0 4{.data=. (ptr+i.len){stream
      cellval=. cellval, null&,@ (":(!.maxpp)) {.fromDouble0 (6+i.8){data
    end.
  case. 16b00fd do. NB. labelss
    rowcolss=. rowcolss, fromWORD0 4{.data=. (ptr+i.len){stream
    cellvalss=. cellvalss, fromDWORD0 (6+i.4){data
  case. 16b0005 do. NB. boolerr biff2
    if. ({.a.)=(8+ptr){stream do. NB. error code ignored
      if. 0=x do.
        rowcol4=. rowcol4, fromWORD0 4{.data=. (ptr+i.len){stream
        cellval4=. cellval4, ({.a.)~:7{data
      else.
        rowcol=. rowcol, fromWORD0 4{.data=. (ptr+i.len){stream
        cellval=. cellval, null&,@ (":(!.maxpp)) ({.a.)~:7{data
      end.
    end.
  case. 16b0205 do. NB. boolerr
    if. ({.a.)=(7+ptr){stream do. NB. error code ignored
      if. 0=x do.
        rowcol4=. rowcol4, fromWORD0 4{.data=. (ptr+i.len){stream
        cellval4=. cellval4, ({.a.)~:6{data
      else.
        rowcol=. rowcol, fromWORD0 4{.data=. (ptr+i.len){stream
        cellval=. cellval, null&,@ (":(!.maxpp)) ({.a.)~:6{data
      end.
    end.
  case. 16b00bd do. NB. multrk
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
  case. 16b0006 do. NB. formula, only cached result read
    data=. (ptr+i.len){stream
    if. 16b200 = biffver do. NB. biff2
      if. (255 255{a.)-:13 14{data do.
        if. (0{a.)=7{data do. NB. string result
          lookrowcol=. fromWORD0 4{.data
          lookstr=. 2
        elseif. (1{a.)=7{data do. NB. boolerr
          if. 0=x do.
            rowcol4=. rowcol4, fromWORD0 4{.data
            cellval4=. cellval4, ({.a.)~:9{data
          else.
            rowcol=. rowcol, fromWORD0 4{.data
            cellval=. cellval, null&,@ (":(!.maxpp)) ({.a.)~:9{data
          end.
        end.
      else.  NB. double
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
        if. (0{a.)=6{data do. NB. string result
          lookrowcol=. fromWORD0 4{.data
          lookstr=. 2
        elseif. (1{a.)=6{data do. NB. boolerr
          if. 0=x do.
            rowcol4=. rowcol4, fromWORD0 4{.data
            cellval4=. cellval4, ({.a.)~:8{data
          else.
            rowcol=. rowcol, fromWORD0 4{.data
            cellval=. cellval, null&,@ (":(!.maxpp)) ({.a.)~:8{data
          end.
NB. blank and multblank records may be ignore
NB.       elseif. (3{a.)=6{data do.  NB. blank
NB.         if. 0=x do.
NB.           rowcolc=. rowcolc, fromWORD0 4{.data
NB.           cellvalc=. cellvalc, null
NB.         else.
NB.           rowcol=. rowcol, fromWORD0 4{.data
NB.           cellval=. cellval, null
NB.         end.
        end.
      else.  NB. double
        if. 0=x do.
          rowcol8=. rowcol8, fromWORD0 4{.data
          cellval8=. cellval8, {.fromDouble0 (6+i.8){data
        else.
          rowcol=. rowcol, fromWORD0 4{.data
          cellval=. cellval, null&,@ (":(!.maxpp)) {.fromDouble0 (6+i.8){data
        end.
      end.
    end.
NB. string for formula result, assume no continue record
  case. 16b0007 do. NB. biff2
    if. 1=lookstr do.
      if. 0=x do.
        rowcolc=. rowcolc, lookrowcol
        cellvalc=. cellvalc, null&,@> {.0 decodestring8 data=. (ptr+i.len){stream
      else.
        rowcol=. rowcol, lookrowcol
        cellval=. cellval, null&,@> {.0 decodestring8 data=. (ptr+i.len){stream
      end.
    end.
NB. string for formula result, assume no continue record
  case. 16b0207 do.
    if. 1=lookstr do.
      if. 16b300 16b400 16b500 e.~ biffver do.  NB. biff3/4/5/7
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
NB. blank and multblank records may be ignore
NB.   case. 16b0201 do. NB. blank
NB.     if. 0=x do.
NB.       rowcolc=. rowcolc, fromWORD0 4{.data=. (ptr+i.len){stream
NB.       cellvalc=. cellvalc, null
NB.     else.
NB.       rowcol=. rowcol, fromWORD0 4{.data=. (ptr+i.len){stream
NB.       cellval=. cellval, null
NB.     end.
NB.   case. 16b00be do. NB. multblank
NB.     rocol=. fromWORD0 4{.data=. (ptr+i.len){stream
NB.     lc=. fromWORD0 _2{.data
NB.     nc=. >:lc-{:rocol
NB.     if. 0=x do.
NB.       rowcolc=. rowcolc, rocol +("1) 0 ,("0) i.nc
NB.       cellvalc=. cellvalc, nc#null
NB.     else.
NB.       rowcol=. rowcol, rocol +("1) 0 ,("0) i.nc
NB.       cellval=. cellval, nc#null
NB.     end.
  case. 16b0004 do. NB. biff2 label, assume no continue record
    if. 0=x do.
      rowcolc=. rowcolc, fromWORD0 4{.data=. (ptr+i.len){stream
      cellvalc=. cellvalc, null&,@> {.0 decodestring8 7}.data
    else.
      rowcol=. rowcol, fromWORD0 4{.data=. (ptr+i.len){stream
      cellval=. cellval, null&,@> {.0 decodestring8 7}.data
    end.
NB. biff8 does not use label record, but excel will read it
  case. 16b0204 do. NB. label, assume no continue record
    if. 16b300 16b400 16b500 e.~ biffver do.  NB. biff3/4/5/7
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
NB. biff8 does not use rstring record, but excel will read it
  case. 16b00d6 do. NB. rstring, assume no continue record
    if. 16b300 16b400 16b500 e.~ biffver do.  NB. biff3/4/5/7
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

NB. decode rk value
getrk=: 3 : 0
y=. y.
if. 0=2 bitand d=. fromDWORD0 2}.y do. NB. double
  bigendian=: ({.a.)={. 1&(3!:4) 1  NB. 0 little endian   1 big endian
  if. 0=bigendian do.
    rk=. fromDouble0 toDWORD0 0, d bitand DFH '0xfffffffc'
  else.
    rk=. fromDouble0 toDWORD0 0,~ d bitand DFH '0xfffffffc'
  end.
else.  NB. integer
  rk=. _2 bitsha d bitand DFH '0xfffffffc'
end.
if. 1 bitand d do.  NB. scale factor
  rk=. rk%100
end.
<{.rk
)

NB. ---------------------------------------------------------
NB. Covers for reading sheet(s) from an Excel workbook
NB. redeveloped by Ric Sherlcok from original by bill lam
NB. ---------------------------------------------------------

NB.*readexcelsheets v Reads one or more sheets from an Excel file
NB. returns: 2-column matrix with a row for each sheet
NB.       0{"1 boxed list of sheet names
NB.       1{"1 boxed list of boxed matrices of sheet contents
NB. y is: 1 or 2-item boxed list:
NB.       0{ filename of Excel workbook
NB.       1{ [default 0] optional switch to return all cells contents as strings
NB. x is: one of [default is 0]:
NB.       * numeric list of indicies of sheets to return
NB.       * boxed list of sheet names to return
NB.       * '' - return all sheets
NB. EG:   0 readexcelsheets 'test.xls'
NB. reads Excel Versions 5, 95, 97, 2000, XP, 2003
NB. biff5  excel 5  biff7 excel 97   biff8 excel 97, xp, 2003
readexcelsheets=: 3 : 0
  y=. y.
  0 readexcelsheets y
:
  y=. y. [ x=. x.
 try.
  'fln strng'=. 2{.!.(<0) boxopen y
  x=. boxopen x
  locs=.'' NB. store locales created
  (msg=.'file not found') assert fexist fln
  locs=.locs,ole=. fln coxnew 'olestorage'
  if. 0=#wks=. getppssearch__ole 'Workbook' ; 1 ; 1 do.              NB. biff8
    if. 0=#wks=. getppssearch__ole 'Book' ; 1 ; 1 do.                NB. biff5/7
      (msg=.'unknown Excel file format') assert 16b40009 16b60209 16b60409 e.~ fromDWORD0 freadx fln;0 4  NB. biff2/3/4
    end.
  end.
  locs=. locs,wks
  locs=.locs,ex=. coxnew 'biffread'
  if. #wks do.
    wk=. {.wks
    0&create__ex data__wk
  NB. get worksheet location
    if. x-:<'' do.
      x=. i.#worksheets__ex
    elseif. -. */ 1 4 8 e.~ 3!:0 every x do.
      x=. x i.~ {."1 worksheets__ex
    elseif. do. x=. >x NB. unbox numeric list
    end.
    (msg=.'worksheet not found') assert x<#worksheets__ex
    'name location'=. |: x{worksheets__ex
  else.
    0&create__ex fread fln
    location=. {.>x
  end.
  'ix cell'=. <"1 |: ,:^:(1:=#@:$) strng&readsheet__ex"0 location NB. read worksheets
  for_l. |.locs do. NB. housekeeping
    destroy__l ''
    locs=. locs -. l
  end. 
  rcs=. 0&>.@(>./ >:@- <./) each ix   NB. transform cell records to matrix
  offset_biffread_=: >{: off=. <./ each ix   NB. store offset in static class variable if needed
  ix=. ix -"1 each off
  m=. 0$0
  for_sht. cell do.
    n=. (sht_index{::rcs)$a:
    n=. (>sht) (<"1^:(0:<#) sht_index{::ix)} n
    m=. m,<n
  end.
  dtb=. #~ ([: +./\. ' '&~:)
  (<@dtb"1 name) ,. m
  NB. convert excel date
  NB. todate (+&36522) 38335 --> 2004 12 14
 catch.
   for_l. |.locs do. NB. housekeeping
     destroy__l '' 
     locs=. locs -. l 
   end.
   smoutput 'readexcelsheets: ',msg
 end.
)

NB.*readexcelsheetsstring v Reads contents of one or more sheets from an Excel file as strings
NB. see readexcelsheets
readexcelsheetsstring=: 3 : 0
  y=. y.
  0 readexcelsheetsstring y
  :
  y=. y. [ x=. x.
  y=. (boxopen y),<1  NB. add string specifier
  x readexcelsheets y
)

NB.*readexcelworkbook v Reads all sheets from an Excel file
NB. see readexcelsheets
readexcelworkbook=: 3 : 0
  y=. y.
  '' readexcelsheets y
)

NB.*readexcel v Reads contents of a sheet from an Excel file
NB. see readexcelsheets
readexcel=: 3 : 0
  y=. y.
  0 readexcel y
  :
  y=. y. [ x=. x.
  x=. {.^:(3!:0~:2:) x NB. ensure single sheet
  ;{:"1 x readexcelsheets y
)

NB.*readexcelstring v Reads contents of a sheet from an Excel file as strings
NB. see readexcelsheets
readexcelstring=: 3 : 0
  y=. y.
  0 readexcelstring y
  :
  y=. y. [ x=. x.
  y=. (boxopen y),<1  NB. add string specifier
  x readexcel y
)

NB. ---------------------------------------------------------
NB.*readexcelsheetnames v Reads sheet names from Excel workbook
NB. returns: boxed list of sheet names
NB. y is: Excel file name
NB. eg: sheetnamesexcel 'test.xls'
NB. read Excel Versions 5, 95, 97, 2000, XP, 2003
NB. biff5  excel 5  biff7 excel 97   biff8 excel 97, xp, 2003
readexcelsheetnames=: 3 : 0
  y=. y. [ x=. x.
 try.
  fln=. boxopen y
  locs=.'' NB. store locales created
  (msg=.'file not found') assert fexist fln
  locs=.locs,ole=. fln conew 'olestorage'
  if. 0=#wks=. getppssearch__ole 'Workbook' ; 1 ; 1 do.              NB. biff8
    if. 0=#wks=. getppssearch__ole 'Book' ; 1 ; 1 do.                NB. biff5/7
      (msg=.'unknown Excel file format') assert 16b40009 16b60209 16b60409 e.~ fromDWORD0 freadx fln;0 4  NB. biff2/3/4
    end.
  end.
  locs=. locs,wks
  locs=.locs,ex=. conew 'biffread'
  if. #wks do.
    wk=. {.wks
    0&create__ex data__wk
  else.
    0&create__ex fread y
  end.
  NB. read worksheet
  nms=. {."1 worksheets__ex
  for_l. |.locs do. destroy__l '' end. NB. housekeeping
  nms
 catch.
   for_l. |.locs do. destroy__l '' end.
   smoutput 'readexcelsheetnames: ',msg
 end.
)

NB. cover function to dump worksheet
NB. sst shared string
NB. (bk)records= type index length
NB. (bk)bytes= byte stream for records
NB. x sheet index or name
NB. y Excel file name
NB. 0 dumpexcel 'test.xls'
dumpexcel=: 0&$: : (4 : 0)
y=. y. [ x=. x.
assert. fexist y
ole=. (>y) coxnew 'olestorage'
if. 0=#wks=. getppssearch__ole 'Workbook' ; 1 ; 1 do.              NB. biff8
  if. 0=#wks=. getppssearch__ole 'Book' ; 1 ; 1 do.                NB. biff5/7
    assert. 16b40009 16b60209 16b60409 e.~ fromDWORD0 freadx y;0 4  NB. biff2/3/4
  end.
end.
ex=. coxnew 'biffread'
if. #wks do.
  wk=. {.wks
  1&create__ex data__wk       NB. 1=debug mode
NB. get worksheet location
  if. 2 131072 e.~ 3!:0 x do. x=. (<x) i.~ {.("1) worksheets__ex end.
  assert. x<#worksheets__ex
  'name location'=. x{worksheets__ex
else.
  1&create__ex fread y
  location=. x
end.
NB. get worksheet location
NB. read worksheet
'ix cell'=. 0&readsheet__ex location
NB. workbook global
sst__=: sst__ex
bkrecords__=: bkrecords__ex
bkbytes__=: bkbytes__ex
NB. worksheet
records__=: records__ex
bytes__=: bytes__ex
NB. housekeeping
destroy__ex ''
for_wk. wks do. destroy__wk '' end.
destroy__ole ''
''
)

NB.  populate z locale
readexcel_z_=: readexcel_biffread_
readexcelstring_z_=: readexcelstring_biffread_
dumpexcel_z_=: dumpexcel_biffread_
readexcelsheetnames_z_=: readexcelsheetnames_biffread_
readexcelsheets_z_=: readexcelsheets_biffread_
readexcelsheetsstring_z_=: readexcelsheetsstring_biffread_
readexcelworkbook_z_=: readexcelworkbook_biffread_
