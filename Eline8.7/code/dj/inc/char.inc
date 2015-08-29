<%
function htmlencode1(fString)
if fString<>"" and not isnull(fstring) then
	fString = replace(fString, "&gt;", ">")
	fString = replace(fString, "&lt;", "<")

	fString = Replace(fString, "&nbsp;", chr(32))
	fString = Replace(fString, "</P><P>", CHR(10) & CHR(10))
	fString = Replace(fString, "<BR>", CHR(10))
	htmlencode1=fString
else
	htmlencode1=""
end if
end function

function htmlencode2(fString)
if fString<>"" and not isnull(fstring) then

	fString = replace(fString, ">", "&gt;")
	fString = replace(fString, "<", "&lt;")

	fString = Replace(fString, chr(32), "&nbsp;")
	fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
	fString = Replace(fString, CHR(10), "<BR>")
	htmlencode2=fString
else
	htmlencode2=""
end if
end function
function htmlencode3(fString)
if fString<>"" and not isnull(fstring) then

	fString = replace(fString, "&gt;", ">")
	fString = replace(fString, "&lt;", "<")

'	fString = Replace(fString, "&nbsp;", chr(32))
'	fString = Replace(fString, "</P><P>", CHR(10) & CHR(10))
'	fString = Replace(fString, "<BR>", CHR(10))
	htmlencode3=fString
else
	htmlencode3=""
end if
end function
function strlen(str)
dim p_len
p_len=0
strlen=0
if trim(str)<>"" then
p_len=len(trim(str))
for xx=1 to p_len
if asc(mid(str,xx,1))<0 then
strlen=int(strlen) + 2
else
strlen=int(strlen) + 1
end if
next
end if
end function

function strvalue(str,lennum)
dim p_num
dim i
if strlen(str)<=lennum then
strvalue=str
else
p_num=0
x=0
do while not p_num > lennum-2
x=x+1
if asc(mid(str,x,1))<0 then
p_num=int(p_num) + 2
else
p_num=int(p_num) + 1
end if
strvalue=left(trim(str),x)&"бн"
loop
end if
end function
function checkStr(str)
	if isnull(str) then
		checkStr = ""
		exit function 
	end if
	checkStr=replace(str,"'","''")
end function
%>
