<!--#include file="conn.asp"-->
<%
dim bbsurl,lockurl
lockurl=""
'ֻ���������ַ,Ҫ��"HTTP://"��ͷ,Ϊ���򲻿��Ŵ˹���.(���������ַ���ƣ�Ҫ��","�ָ���)
'����ֻ�����������ַ����: lockurl="http://www.artistsky.net/,http://www.artbbs.net/"
bbsurl="http://zhzx.jjedu.org/eline/bbs/"  '����д����̳����ȷ��ַ,Ҫ��"HTTP://"��ͷ

'*************************************
'�ϴ�����CONN.ASPͬ����Ŀ¼��
'���ϵ�ַ����һ��Ҫ�޸�,���������õ�������ȥ�����ϵ���̳.
'��������,��������һ���ϴ���newscode.ASP�ļ����е���(newscode.ASP����ǰҪ�޸ĵ��ò���)
'	FSSUNWIN	2003.10.1
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

dim rs,sql,i,n
	if trim(request("n"))<>"" and IsNumeric(request("n")) then
	n=cint(request("n"))
	else
	n=1
	end if
	select case cint(request("orders"))
	case 1
		call tongji()
	case 2
		call topuser()
	case 3
		call adduser()
	case 4
		if trim(request("boardid"))<>"" and isnumeric(trim(request("boardid"))) then
			if trim(request("boardid"))=0 then
			call board(0,cint(trim(request("stats"))),cint(trim(request("model"))))
			else
			call board(cint(trim(request("boardid"))),cint(trim(request("stats"))),cint(trim(request("model"))))
			end if
		else
			call board("all",cint(trim(request("stats"))),cint(trim(request("model"))))
		end if
	case 5
		call bbsnews()
	case 6
		call bbslinks()
	case else
		call tongji()
	end select
	CloseDatabase
            
function allboys()
dim tmprs
    tmprs=conn.execute("Select count(userid) from [user] where sex='1' ")  
    allboys=tmprs(0)  
set tmprs=nothing  
if isnull(allboys) then allboys=0  
end function

function allgirls()  
dim tmprs
	tmprs=conn.execute("Select count(userid) from [user] where sex='0' ")  
    allgirls=tmprs(0)  
set tmprs=nothing  
if isnull(allgirls) then allgirls=0  
end function
function online()
dim tmprs
     tmprs=conn.execute("Select count(id) from online")  
     online=tmprs(0)  
set tmprs=nothing  
if isnull(online) then online=0 
end function
sub tongji()
dim bbsinfo
sql="select top 1 TopicNum,BbsNum,TodayNum,UserNum,lastUser,yesterdaynum,maxpostnum,maxpostdate from config where active=1"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if not rs.eof then
	bbsinfo = rs.GetString(,,"|||","$$$"," ")
	end if
	rs.close
	set rs=nothing
	bbsinfo=split(bbsinfo,"|||")
%> 
 document.write('��- �������� <%=bbsinfo(0)%><br> ��- ��̳���� <%=bbsinfo(1)%><br> ��- ע������ <%=bbsinfo(3)%><br> ��- ��̳���� <font color=red><%=online()%></font><br> ��- �½���Ա <font color=red><%=bbsinfo(4)%></font><br> ��- ��̳Ů�� <%=allgirls()%> λ<br> ��- ��̳���� <%=allboys()%> λ');
 document.write('<br> ��- �������� <font color=red><%=bbsinfo(2)%></font><br> ��- �������� <%=bbsinfo(5)%><br> ��- �߷����� <%=bbsinfo(6)%>')
<% end sub

sub topuser()
	set rs=server.createobject("adodb.recordset")
	sql="select top "&n&" userid,username from [user] order by article desc,userid desc"
	rs.open sql,conn,1,1
	If Not RS.Eof then
		SQL=Rs.GetRows(-1)
    end if
    rs.close:set rs=nothing
	For i=0 To Ubound(SQL,2)
	response.write "document.write('<font face=Wingdings color=#FFAA39>J</font> <a href="&bbsurl&"dispuser.asp?id="& SQL(0,i) &" target=_blank title=�鿴"&SQL(1,i)&"�ĸ�������>');document.write('"
    response.write SQL(1,i)
	response.write "</a>');document.write('<br>');"
	next
end sub

sub adduser()
	set rs=server.createobject("adodb.recordset")
	sql="select top "&n&" userid,username from [user] order by userid desc"
	rs.open sql,conn,1,1
	If Not RS.Eof then
		SQL=Rs.GetRows(-1)
    end if
    rs.close:set rs=nothing
	For i=0 To Ubound(SQL,2)
	response.write "document.write('<font face=Wingdings color=#FFAA39>J</font> <a href="&bbsurl&"dispuser.asp?id="& SQL(0,i) &" target=_blank title=�鿴"&SQL(1,i)&"�ĸ�������>');document.write('"
    response.write SQL(1,i)
	response.write "</a>');document.write('<br>');"
	next
end sub

sub board(id,stat,model)
dim sqlstr,k,ii,jump,t,tbr
ii=0
t=0
tbr=5
sqlstr=""
if trim(request("depth"))<>"" and isnumeric(trim(request("depth"))) then
sqlstr=" and depth<="&cint(trim(request("depth")))
end if
if id="all" then
	if trim(request("depth"))<>"" then sqlstr="where "&replace(sqlstr," and","",1,1)
	sql="select boardid,boardType,depth,Board_Setting from board "&sqlstr&" order by rootid,orders "
else
	sql="select boardid,boardType,depth,Board_Setting from board where ParentID="&cint(id)&" "&sqlstr&" order by rootid,orders "
end if
set rs=conn.execute(sql)
If Not RS.Eof then
		SQL=Rs.GetRows(-1)
end if
rs.close:set rs=nothing
For k=0 To Ubound(SQL,2)
if cint(stat)=1 and mid(SQL(3,k), 3, 1)=1 then
jump=true
else
jump=false
end if
if jump=false then
	if model=1 then
		select case SQL(2,k)
		case 0
		response.write "document.write('<br><font face=Wingdings color=#FFAA39><b>O</b></font> ');"
		ii=0
		t=0
		case 1
		t=t+1
		if ii=0 then
			response.write "document.write('<BR>&nbsp;&nbsp;');"
			ii=1
			t=1
		else
			if t>tbr then
			response.write "document.write('<BR>&nbsp;&nbsp;');"
			t=1
			else
			response.write "document.write('--');"
			end if
		end if
		end select
		if SQL(2,k)<2 then
		response.write "document.write('<a href="&bbsurl&"list.asp?boardid="& SQL(0,k) &" target=_blank title=""��ӭ�ι�"& SQL(1,k) &"!"">');"
		response.write "document.write('"& SQL(1,k) &"');"
		response.write "document.write('</a>');"
		end if
	else
		select case SQL(2,k)
		case 0
		response.write "document.write('��');"
		ii=ii+1
		response.write "document.write('("&ii&")');"
		case 1
		response.write "document.write('&nbsp;&nbsp;<font color=cc0000><b>��</b></font>');"
		end select
		if SQL(2,k)>1 then
		for i=2 to SQL(2,k)
		response.write "document.write('&nbsp;&nbsp;��');"
		next
		response.write "document.write('&nbsp;&nbsp;��');"
		end if
		response.write "document.write('<a href="&bbsurl&"list.asp?boardid="& SQL(0,k) &" target=_blank title=""��ӭ�ι�"& SQL(1,k) &"!"">');"
		response.write "document.write('"& SQL(1,k) &"');"
		response.write "document.write('</a><br>');"
	end if
end if
next
end sub

sub bbsnews()
dim sqlstr,k
if trim(request("boardid"))<>"" and isnumeric(trim(request("boardid"))) then
sqlstr=" where boardid="&cint(trim(request("boardid")))&" "
else
sqlstr=""
end if
sql="select top "&n&" boardid,title,username,addtime,id from bbsnews "&sqlstr&" order by id desc"
set rs=conn.execute(sql)
If Not RS.Eof then
	SQL=Rs.GetRows(-1)
	else
	response.write "document.write('<b><a href="&bbsurl&"announcements.asp?boardid=0 target=_blank><ACRONYM TITLE=""��ǰû�й���"">��ǰû�й���</ACRONYM></a></b> ("& now() &")');"
	exit sub
end if
rs.close:set rs=nothing
select case cint(request("model"))
case 1
	For k=0 To Ubound(SQL,2)
	response.write "document.write('<font face=Wingdings color=#FFAA39>w</font>&nbsp;&nbsp;<a href="""&bbsurl&"announcements.asp?action=showone&boardid="& SQL(0,k) &"&id="&SQL(4,k)&""" target=""_blank"" title=""������:"&SQL(2,k)&"&nbsp;&nbsp;ʱ��:"&SQL(3,k)&""">');"
	if request("tlen")<>"" and isnumeric(trim(request("tlen"))) then
		if len(SQL(1,k))>Cint(request("tlen")) then
		response.write "document.write('"&left(SQL(1,k),request("tlen"))&"...');"
		else
		response.write "document.write('"&SQL(1,k)&"');"
		end if
	else
	response.write "document.write('"&SQL(1,k)&"');"
	end if
	response.write "document.write('</a><br>');"
	next
case 2
	response.write "document.write('<marquee id=""shownews"" behavior=""alternate"" direction=""left"" scrollamount=""4"" scrolldelay=""1"" hspace=""0"" vspace=""0"">');"
	For k=0 To Ubound(SQL,2)
	response.write "document.write('<font face=Wingdings color=#FFAA39>w</font>&nbsp;&nbsp;<a href="""&bbsurl&"announcements.asp?action=showone&boardid="& SQL(0,k) &"&id="&SQL(4,k)&""" target=""_blank"" title=""������:"&SQL(2,k)&"&nbsp;&nbsp;ʱ��:"&SQL(3,k)&""" onmouseover=""document.all.shownews.stop();"" onmouseout=""document.all.shownews.start();"">');"
	response.write "document.write('"&SQL(1,k)&"');"
	response.write "document.write('</a>&nbsp;&nbsp;&nbsp;&nbsp;');"
	next
	response.write "document.write('</marquee>');"
case else
	exit sub
end select
end sub

sub bbslinks()

end sub

'bbsurl=getservepath(request.ServerVariables("server_name")&request.ServerVariables("URL"))
'function getservepath(str)
'dim tmpstr
'tmpstr=split(str,"/")
'getservepath="http://"&replace(str, tmpstr(ubound(tmpstr)), "")
'end function
%>