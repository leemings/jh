<!--#include file="conn.asp"-->
<%
'DVBBS 6.0 ������̳��ҳ����-----��̳չ������
'############����Ϊ�޸���######################
dim lockurl
lockurl=""
'ֻ���������ַ,Ҫ��"HTTP://"��ͷ,Ϊ���򲻿��Ŵ˹���.(���������ַ���ƣ�Ҫ��","�ָ���)
'����ֻ�����������ַ����: lockurl="http://www.artistsky.net/,http://www.artbbs.net/"
Const IsSqlDataBase=0						'�������ݿ����1ΪSQL���ݿ⣬0ΪAccess���ݿ�
Const bbsurl="http://www.artbbs.net/"		'����д����̳����ȷ��ַ,Ҫ��"HTTP://"��ͷ
Const lockboardid="1,2"						'����д���Ƶ��õ���̳���ID,�ö��Ÿ���������boardid=allʱ����Ч)
const picwidth="110"						'�����ļ����
const picheight="110"						'�����ļ��߶�
const tdcolor1="#FFFFEE"					'��������ɫ
const tdcolor2="#FFE9BB"					'��������ɫ
'############����Ϊ�޸���######################
'*************************************
'�ϴ�����CONN.ASPͬ����Ŀ¼��
'���ϵ�ַ����һ��Ҫ�޸�,���������õ�������ȥ�����ϵ���̳.
'��������,��������һ���ϴ���newscode.ASP�ļ����е���(newscode.ASP����ǰҪ�޸ĵ��ò���)
'	FSSUNWIN	2003.10.20
'*************************************
dim n,i,t,F_Filename,tab,orders,tdcolor

if trim(lockurl)<>"" and checkserver(lockurl)=false then
	response.write "document.write ('���ݱ�����,��ֹ������վ�����!');"
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
	'F_Typ : 1=ͼƬ��,2=FLASH��,3=���ּ�,4=��Ӱ��,0=�ļ���
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
	if instr(SQL(3,i),"/")=0 then '�ж��ļ��Ƿ���̳������������ñ��еļ�¼��
	F_Filename=bbsurl&"UploadFile/"&SQL(4,i)
	end if
	response.write "document.write('<td align=""center"" style=""background-color:"&tdcolor&""">');"
	response.write "document.write('<A HREF="""&bbsurl&"fileshow.asp?boardid="&SQL(1,i)&"&id="&SQL(0,i)&""" title=""��&nbsp;&nbsp;&nbsp;&nbsp;��:"&htmlencode(SQL(5,i))&"&#13;&#10;�ṩ��:"&htmlencode(SQL(3,i))&"&#13;&#10;�����:"&SQL(2,i)&" ��&#13;&#10;ʱ&nbsp;&nbsp;&nbsp;&nbsp;��:"&SQL(8,i)&""" target=""_blank"">');"
	if SQL(6,i)=1 then
	response.write "document.write('<IMG SRC="""&htmlencode(F_Filename)&""" style=""border: 1 solid #000000"" width="&picwidth&" height="&picheight&" >');"
	else
	response.write "document.write('"&SQL(7,i)&" ���ļ�');"
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

Rem ����HTML����
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
	HTMLEncode="δ��д��"
end if
end function
%>

