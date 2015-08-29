<!--#include file="conn.asp"-->
<%
'DVBBS 6.0 动网论坛首页调用-----论坛展区调用
'############以下为修改项######################
dim lockurl
lockurl=""
'只允许调用网址,要以"HTTP://"开头,为空则不开放此功能.(可允许多网址限制，要以","分隔。)
'例如只允许此两个网址调用: lockurl="http://www.artistsky.net/,http://www.artbbs.net/"
Const IsSqlDataBase=0						'定义数据库类别，1为SQL数据库，0为Access数据库
Const bbsurl="http://www.artbbs.net/"		'请填写你论坛的正确地址,要以"HTTP://"开头
Const lockboardid="1,2"						'请填写定制调用的论坛版块ID,用逗号隔开。（当boardid=all时才生效)
const picwidth="110"						'定义文件宽度
const picheight="110"						'定义文件高度
const tdcolor1="#FFFFEE"					'定义表格１颜色
const tdcolor2="#FFE9BB"					'定义表格２颜色
'############以上为修改项######################
'*************************************
'上传到与CONN.ASP同级的目录下
'以上地址参数一定要修改,否则所调用的链接是去了以上的论坛.
'若有问题,可以运行一起上传的newscode.ASP文件进行调试(newscode.ASP运行前要修改调用参数)
'	FSSUNWIN	2003.10.20
'*************************************
dim n,i,t,F_Filename,tab,orders,tdcolor

if trim(lockurl)<>"" and checkserver(lockurl)=false then
	response.write "document.write ('数据被保护,禁止被其他站点调用!');"
	response.end	
end if

Private function checkserver(str)
	dim i,servername
	checkserver=false
	if str="" then exit function
	str=split(Cstr(str),",")
	servername=Request.ServerVariables("HTTP_REFERER")
	for i=0 to Ubound(str)
	if right(str(i),1)="/" then str(i)=left(trim(str(i)),len(str(i))-1)
		if Lcase(left(servername,len(str(i))))=Lcase(str(i)) then
			checkserver=true
			exit for
		else
			checkserver=false
		end if
	next
end function

dim SqlNowString,rs,sql,searchstr
if IsSqlDataBase=0 then
SqlNowString="Now()"
else
SqlNowString="GetDate()"
end if

t=0:tab=4:n=1
orders="F_AddTime"
tdcolor=tdcolor1
	if trim(request("n"))<>"" and IsNumeric(request("n")) then
	n=cint(request("n"))
	end if
	if request("tab")<>"" and isnumeric(request("tab")) then
	tab=cint(request("tab"))
	end if
	if trim(request("orders"))=1 then
		orders="F_ViewNum"
	end if
	'F_Typ : 1=图片集,2=FLASH集,3=音乐集,4=电影集,0=文件集
	if trim(request("type"))<>"" and IsNumeric(trim(request("type"))) then
		if searchstr="" then
		searchstr="where F_Type="&cint(trim(request("type")))
		else
		searchstr=searchstr&" and F_Type="&cint(trim(request("type")))
		end if
	end if
	if trim(request("boardid"))<>"" and IsNumeric(trim(request("boardid"))) then
		if searchstr="" then
		searchstr="where F_BoardID = "&cint(trim(request("boardid")))
		else
		searchstr=searchstr&" and F_BoardID = "&cint(trim(request("boardid")))
		end if
	else
		if trim(request("lock"))=1 then
			if searchstr="" then
			searchstr="where F_BoardID not in ("&lockboardid&") "
			else
			searchstr=searchstr&" and F_BoardID not in ("&lockboardid&") "
			end if
		elseif trim(request("lock"))=2 then
			if searchstr="" then
			searchstr="where F_BoardID in ("&lockboardid&") "
			else
			searchstr=searchstr&" and F_BoardID in ("&lockboardid&") "
			end if
		end if
	end if
select case request("mode")
case 0
case 1
response.write "document.write('<marquee onMouseOver=this.stop() onMouseOut=this.start() border=0 align=middle scrollamount=1 direction=up scrolldelay=60 behavior=scroll width=100% height=""100%"">');"
case 2
response.write "document.write('<marquee onMouseOver=this.stop() onMouseOut=this.start() border=0 align=middle scrollamount=1 direction=down scrolldelay=60 behavior=scroll width=100% height=""100%"">');"
case 3
response.write "document.write('<marquee onMouseOver=this.stop() onMouseOut=this.start() border=0 align=middle scrollamount=1 behavior=""alternate"" direction=left scrolldelay=60 behavior=scroll width=100% height=""100%"">');"
case 4
response.write "document.write('<marquee onMouseOver=this.stop() onMouseOut=this.start() border=0 align=middle scrollamount=1 behavior=""alternate"" direction=right scrolldelay=60 behavior=scroll width=100% height=""100%"">');"
end select
response.write "document.write('<table border=""0"" cellspacing=""1"" cellpadding=""3"" align=center width=""100%""><tr>');"
	set rs=conn.execute("select top "&n&" F_ID,F_BoardID,F_ViewNum,F_Username,F_Filename,F_Readme,F_Type,F_FileType,F_AddTime from DV_Upfile "&searchstr&" order by "&orders&" desc, F_ID desc")
	If Not RS.Eof then
		SQL=rs.GetRows(-1)
    end if
	rs.close:set rs=nothing
	CloseDatabase

For i=0 To Ubound(SQL,2)
	if tdcolor=tdcolor2 then 
	tdcolor=tdcolor1
	else
	tdcolor=tdcolor2
	end if
	if instr(SQL(3,i),"/")=0 then '判断文件是否本论坛，若不是则采用表中的记录．
	F_Filename=bbsurl&"UploadFile/"&SQL(4,i)
	end if
	response.write "document.write('<td align=""center"" style=""background-color:"&tdcolor&""">');"
	response.write "document.write('<A HREF="""&bbsurl&"fileshow.asp?boardid="&SQL(1,i)&"&id="&SQL(0,i)&""" title=""主&nbsp;&nbsp;&nbsp;&nbsp;题:"&htmlencode(SQL(5,i))&"&#13;&#10;提供者:"&htmlencode(SQL(3,i))&"&#13;&#10;被浏览:"&SQL(2,i)&" 次&#13;&#10;时&nbsp;&nbsp;&nbsp;&nbsp;间:"&SQL(8,i)&""" target=""_blank"">');"
	if SQL(6,i)=1 then
	response.write "document.write('<IMG SRC="""&htmlencode(F_Filename)&""" style=""border: 1 solid #000000"" width="&picwidth&" height="&picheight&" >');"
	else
	response.write "document.write('"&SQL(7,i)&" 类文件');"
	end if
	if request("topic")=1 then response.write "document.write('<br>"&htmlencode(SQL(5,i))&"');"
	response.write "document.write('</A></td>');"
	if t=tab-1 or tab=1 then 
	response.write "document.write('</tr><tr>');"
	end if
	if t>tab-1 then 
		t=1
	else
		t=t+1
	end if
NEXT
response.write "document.write('</tr></table>');"
if request("mode")<>"" and request("mode")<>0 then response.write "document.write('</marquee>');"

Rem 过滤HTML代码
function HTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")
    fString = Replace(fString, CHR(32), "&nbsp;")
    fString = Replace(fString, CHR(9), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "&nbsp; ")
    fString = Replace(fString, CHR(10), "&nbsp; ")
    HTMLEncode = fString
	ELSE
	HTMLEncode="未填写！"
end if
end function
%>

