<!--#include file="conn.asp"-->
<%
'############����Ϊ�޸���######################
dim lockurl
lockurl=""
'ֻ���������ַ,Ҫ��"HTTP://"��ͷ,Ϊ���򲻿��Ŵ˹���.(���������ַ���ƣ�Ҫ��","�ָ���)
'����ֻ�����������ַ����: lockurl="http://www.artistsky.net/,http://www.artbbs.net/"
Const IsSqlDataBase=0						'�������ݿ����1ΪSQL���ݿ⣬0ΪAccess���ݿ�
Const bbsurl="http://www.artbbs.net/"		'����д����̳����ȷ��ַ,Ҫ��"HTTP://"��ͷ
Const lockboardid="1,2,3"				    '����д�������õ���̳���ID,�ö��Ÿ���������LOCK����=1��boardid=allʱ����Ч)
Const picurl="http://www.artbbs.net/face/"	'����ͼ��Ŀ¼��ַ
'############����Ϊ�޸���######################
'bbsurl=getservepath(request.ServerVariables("server_name")&request.ServerVariables("URL"))
'function getservepath(str)
'dim tmpstr
'tmpstr=split(str,"/")
'getservepath="http://"&replace(str, tmpstr(ubound(tmpstr)), "")
'end function
'*************************************
'�ϴ�����CONN.ASPͬ����Ŀ¼��
'���ϵ�ַ����һ��Ҫ�޸�,���������õ�������ȥ�����ϵ���̳.
'��������,��������һ���ϴ���newscode.ASP�ļ����е���(newscode.ASP����ǰҪ�޸ĵ��ò���)
'	FSSUNWIN	2003.10.20
'*************************************
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

dim SqlNowString
if IsSqlDataBase=0 then
SqlNowString="Now()"
else
SqlNowString="GetDate()"
end if
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
end if
end function

	dim rs,sql
	dim orders,reply,topic,isbest
    dim board
    dim bname,ars
    dim postinfo,postname,POSTTIME
	dim NowUseBbs,boardname
	dim i,k,n,sdate,searchdate
	i=0:k=0
	if trim(request("n"))<>"" and IsNumeric(request("n")) then
	n=cint(request("n"))
	else
	n=1
	end if

	if trim(request("sdate"))<>"" and IsNumeric(request("sdate")) then
		sdate=cint(request("sdate"))
		if IsSqlDataBase=1 Then
		searchdate=" and datediff(day,dateandtime,"&SqlNowString&")<"&sdate
		else
		searchdate=" and datediff('d',dateandtime,"&SqlNowString&")<"&sdate
		end if
	else
	searchdate=""
	end if

	if trim(request("orders"))=1 then
		orders="hits"
	else
		orders="dateandtime"
	end if

	if trim(request("boardid"))<>"all" and isnumeric(request("boardid")) then
		board=" and boardid="&cint(trim(request("boardid")))&""
	end if
	if cint(trim(request("lock")))=1 then
		board=board+" and boardid not in ("&lockboardid&") "
	end if

	if cint(request("action"))>2 then
	set rs=conn.execute("select top 1 NowUseBBS from config where active=1 ")
	NowUseBBS=rs(0)
	rs.close
	end if
	if cint(request("action"))=1 then
		'��ʾ����
		if trim(request("orders"))=2 then orders="lastposttime"
		set rs=conn.execute("select top "&cint(n)&" PostUserName,Title,topicid,boardid,dateandtime,topicid,hits,Expression,LastPost from topic where locktopic<>2 "&searchdate&board&" ORDER BY "&orders&" desc,topicid desc")
	elseif cint(request("action"))=2 then
		'��ʾ��������
		if searchdate<>"" then searchdate=replace(searchdate," and"," where")
		if searchdate="" and board<>"" then board=replace(board," and"," where")
		'response.write searchdate&board
		set rs=conn.execute("select top "&cint(n)&" PostUserName,Title,rootid,boardid,dateandtime,Announceid,id,Expression from BestTopic  "&searchdate&board&"  ORDER BY "&orders&" desc")
    else
	    '��ʾ�����ظ�
		set rs=conn.execute("select top "&cint(n)&" username,topic,rootid,boardid,dateandtime,announceid,body,Expression from "&NowUseBbs&" where locktopic<>2 "&searchdate&board&" ORDER BY "&orders&" desc")
	end if
	If Not RS.Eof then
		SQL=Rs.GetRows(-1)
		else
		response.write "document.write('��δ�������ӣ�');"
		response.end
    end if
    rs.close
	set rs=nothing
	For i=0 To Ubound(SQL,2)
		topic=SQL(1,i)
		if topic="" then
			topic=SQL(6,i)
		end if
		if len(topic)>Cint(request("tlen")) then
			topic=left(topic,request("tlen"))&"..."
		end if
		postname=SQL(0,i)
		POSTTIME=SQL(4,i)
		if request("action")=1 and request("reply")=1 then
			if  SQL(8,i)<>"" then
			postinfo=split(SQL(8,i),"$")
			postname=postinfo(0)
			POSTTIME=postinfo(2)
			end if
		end if
		if request("showpic")=1 then
		response.write "document.write('<IMG SRC="""&picurl&SQL(7,i)&""" BORDER=0> ');"
		else
		response.write "document.write('<font face=wingdings color=#FFAA39><B>1</B></FONT> ');"
		end if
		if request("bname")=1 then
			set ars=conn.execute("select BoardType from board where boardid="&SQL(3,i))
				boardname=ars(0)
				ars.close
				response.write "document.write('<a href="&bbsurl&"list.asp?boardid="&SQL(3,i)&" target=_top>"&htmlencode(boardname)&"</a>');"
		end if
		response.write "document.write('<a href="&bbsurl&"dispbbs.asp?boardid="&SQL(3,i)&"&ID="&SQL(2,i)&"&replyID="&SQL(5,i)&" target=_top title="&htmlencode(topic)&">');"
		response.write "document.write('"&htmlencode(topic)&"');"
		response.write "document.write('</a>');"
		select case cint(request("info"))
		case 0
		case 1
		response.write "document.write('��(<a href="&bbsurl&"dispuser.asp?name="&postname&" target=_blank>"&postname&"</a>,<font color=green>"&formatdatetime(POSTTIME,0)&"</font>)');"
		case 2
		response.write "document.write('��(<font color=green>"&POSTTIME&"</font>)');"
		case 3
		response.write "document.write('��(<a href="&bbsurl&"dispuser.asp?name="&postname&" target=_blank>"&postname&"</a>)');"
		case 4
		response.write "document.write('��(<a href="&bbsurl&"dispuser.asp?name="&postname&" target=_blank>"&postname&"</a>');"
		if cint(request("action"))=1 then response.write "document.write(',<font color=green>"&SQL(6,i)&"</font>)');"
		case 5
		if cint(request("action"))=1 then
		response.write "document.write('(<font color=green>"&SQL(6,i)&"</font>)');"
		end if
		case 6
		response.write "document.write('��(<a href="&bbsurl&"dispuser.asp?name="&postname&" target=_blank>"&postname&"</a>,<font color=green>"&formatdatetime(POSTTIME,1)&"</font>)');"
		case 7
		response.write "document.write('��(<font color=green>"&formatdatetime(POSTTIME,1)&"</font>)');"
		case else

		end select
		response.write "document.write('<br>');"
		k=k+1
	Next
	CloseDatabase
%>

