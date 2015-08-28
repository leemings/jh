<%
Response.Expires=0
username=Session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"

'声明connection

dbname="driver={Microsoft Access Driver (*.mdb)};DBQ="&Server.MapPath("../../21ex/setup.asp")
set conn=server.createobject("adodb.connection")
conn.open dbname
set rs=server.createobject("adodb.recordset")


'----------------------------------------------------------------------

'-----------------------------------------------------------------------
'得到文件扩展名的函数
'调用：getfileextname(文件全名)

function getFileExtName(fileName)
dim pos
pos=instrrev(filename,".")
if pos>0 then 
getFileExtName=mid(fileName,pos+1)
else
getFileExtName=""
end if
end function 
'判断空字符串，如是，则用指定值填充
'调用：cnone(字符串,填充字符串)

function cnone(panduan,zhiding)
if trim(panduan)="" then
cnone=zhiding
else
cnone=panduan
end if
end function
'转换日期型或字符日期型为全格式如：2000-01-01
'调用：ccdate(日期变量或日期格式字符串)

function ccdate(sdate)
temp=cdate(sdate)
if len(month(temp))=1 then
m="0"&month(temp)
else
m=month(temp)
end if

if len(day(temp))=1 then
d="0"&day(temp)
else
d=day(temp)
end if
ccdate=year(temp)&"-"&m&"-"&d
end function
'转换时间字符串为全格式,如：01:01:01
'调用：cctime(字符串)

function cctime (stime)
if len(hour(stime))=1 then
h="0"&hour(stime)
else
h=hour(stime)
end if
if len(minute(stime))=1 then
m="0"&minute(stime)
else
m=minute(stime)
end if
if len(second(stime))=1 then
s="0"&second(stime)
else
s=second(stime)
end if
cctime=h&":"&m&":"&s
end function
'转换为写入数据库的格式，去掉非法字符
'调用：putmeg(字符串)

function putmeg (inputmsg)
putmeg=replace(putmeg,"<",".")
putmeg=replace(inputmsg,chr(13)&chr(10),"<br>")
putmeg=replace(putmeg," ","&nbsp;&nbsp;")
putmeg=replace(putmeg,"'","’")


end function
'按概率返回true，精确至0.1%
'调用：prob(概率)
function prob(temprob)
randomize
if temprob>100 then
Response.Write( "Flush内部函数 prob 错误！<br>")
Response.Write("超出概率范围")
Response.End
end if
pnumber=(rnd*1000)/10
if pnumber<temprob then
prob=true
else
prob=false
end if
end function
'还原数据库写入格式（putmsg的反操作）
'调用：getmsg（字符串或数据库字段值）

function getmsg(inputmsg)
getmsg=replace(inputmsg,"’","'")
getmsg=replace(getmsg,"&nbsp;&nbsp;"," ")
getmsg=replace(getmsg,"<br>",chr(13)&chr(10))
end function
'判断email地址的正确性
'调用：isemail(email地址字符串)，如正确返回true,不正确则返回false

Function IsEmail(Email)
ValidFlag = False
If (Email <> "") And (InStr(1, Email, "@") > 0) And (InStr(1, Email, ".") > 0) Then
atCount = 0
SpecialFlag = False
For atLoop = 1 To Len(Email)
atChr = Mid(Email, atLoop, 1)
If atChr = "@" Then atCount = atCount + 1
If (atChr >= Chr(32)) And (atChr <= Chr(44)) Then SpecialFlag = True
If (atChr = Chr(47)) Or (atChr = Chr(96)) Or (atChr >= Chr(123)) Then SpecialFlag = True
If (atChr >= Chr(58)) And (atChr <= Chr(63)) Then SpecialFlag = True
If (atChr >= Chr(91)) And (atChr <= Chr(94)) Then SpecialFlag = True
Next
If (atCount = 1) And (SpecialFlag = False) Then
BadFlag = False
tAry1 = Split(Email, "@")
UserName = tAry1(0)
DomainName = tAry1(1)
If (UserName = "") Or (DomainName = "") Then BadFlag = True
If Mid(DomainName, 1, 1) = "." then BadFlag = True
If Mid(DomainName, Len(DomainName), 1) = "." then BadFlag = True
ValidFlag = True
End If
End If
If BadFlag = True Then ValidFlag = False
IsEmail = ValidFlag
End Function
'=====================================================

'截取一定长度的字符串，其余补" ..."
'调用：cutstr(字符串，截取长度)

function cutstr(tempstr,tempwid)


if instr(tempstr,"&nbsp;") then
tempstr=replace(tempstr,"&nbsp;"," ")
end if
if instr(tempstr,"<br>") then
tempstr=replace(tempstr,"<br>"," ")
end if

if len(tempstr)>tempwid then
cutstr=left(tempstr,tempwid)&" ..."
else
cutstr=tempstr
end if
end function
'=====================================================

%>
