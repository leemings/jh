<%Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
%>
<!--#include file="set.asp"-->
<%
bid=Request.QueryString("bid")
username=session("Ba_jxqy_username")
usergrade=session("Ba_jxqy_usergrade")
usercorp=session("Ba_jxqy_usercorp")
if username="" then Response.Redirect "../error.asp?id=016"
nowtime=now()
if not isnumeric(bid) or bid="" then Response.Redirect "../error.asp?id=220"
activepage=Request("activepage")
if activepage="" or not isnumeric(activepage) then activepage=1
activepage=clng(activepage)
if activepage<1 then activepage=1
msg="<html><head><meta http-equiv='Content-Type' content='text/html; charstet=gb2312'><title>"&Ba_pagename&"</title><link rel='stylesheet' href='../chatroom/css.css'></head><body bgcolor="&Ba_bgcolor&" background="&Ba_bgimage&" oncontextmenu='self.event.returnValue=false' topmargin=0 leftmargin=0 >"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from ��� where id="&bid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=220"
msg=msg&"<table border=0 width=100% bgcolor=cccccc><tr><td>��̳���ƣ�"&showcname(rst("id"),rst("����"),rst("˵��"))&"</td><td>����:"
cname=rst("����")
host=split(rst("����"),";")
for i=0 to ubound(host)
	msg=msg&showun(host(i))
next
msg=msg&"</td><td width='30%' ><marquee>"&rst("˵��")&"</marquee></td><td align=right>"&showhomepage(Ba_homepage,Ba_pagename)&"</td></tr></table>"
rst.Close
msg=msg&"<table border=0 bgcolor='"&Ba_bgcolor&"' background='"&Ba_bgimage&"' width='100%'><tr><td><input type=button onclick="&chr(34)&"javascript:location.replace('addnewbbs.asp?bid="&bid&"');"&chr(34)&"  value=' �� �� '></td><td><font color=00ff00>"&username&"</font>��ӭ���Ĺ���</td><td align=right><a href='today.asp?bid="&bid&"' onmouseover="&chr(34)&"window.status='�鿴��������';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title="&chr(34)&"�鿴��������"&chr(34)&">��������</a>>>������</td></tr></table><table border=1 width=90% align=center bgcolor=cccccc><tr bgcolor=fffe48 align=center><td>&nbsp;</td><td>&nbsp;</td><td>���±���</td><td>�´���</td><td>����</td><td>��[��]</td><td>���ظ�</td><td>������ʱ��</td><td>���</td></tr>"
rst.Open "select * from bbs where ��̳="&bid&" and ����=0 and ɾ��=false and ����=True order by �ظ�ʱ�� desc",conn,1,3
if not(rst.EOF or rst.BOF) then
rst.PageSize=Ba_pagesize
if activepage>rst.PageCount then activepage=rst.PageCount
rst.AbsolutePage=activepage
bgc="bgcolor=f7f7f7"
for i=1 to rst.PageSize
	if rst.EOF or rst.BOF then exit for
	if bgc="bgcolor=f7f7f7" then 
		bgc=""
	else
		bgc="bgcolor=f7f7f7"
	end if	
	msg=msg&"<tr "&bgc&"><td>"
	if datediff("d",rst("�ظ�ʱ��"),nowtime)>=1 and rst("�ظ�")<Ba_hotbbsnum then
		msg=msg&"<img src='../images/closed.gif'></td><td>"
	elseif	datediff("d",rst("�ظ�ʱ��"),nowtime)>=1 and rst("�ظ�")>=Ba_hotbbsnum then
		msg=msg&"<img src='../images/hotclosed.gif'></td><td>"
	elseif	datediff("d",rst("�ظ�ʱ��"),nowtime)<1 and rst("�ظ�")<Ba_hotbbsnum then
		msg=msg&"<img src='../images/closedb.gif'></td><td>"
	else
		msg=msg&"<img src='../images/hotclosedb.gif'></td><td>"
	end if	
	if rst("����")=True then
		msg=msg&"<img src='../images/essence.gif'></td><td><a href='showbbs.asp?bid="&bid&"&fid="&rst("ID")&"' onmouseover="&chr(34)&"window.status='�鿴��������ػظ���';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"
	else
		msg=msg&"&nbsp;</td><td><a href='showbbs.asp?bid="&bid&"&fid="&rst("ID")&"'  onmouseover="&chr(34)&"window.status='�鿴��������ػظ���';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"
	end if
	if len(rst("����"))<15 and datediff("h",rst("�ظ�ʱ��"),nowtime)<1 then 
		msg=msg&rst("����")&"</a><img src='../images/new.gif'></td><td align=center>"
	elseif len(rst("����"))<15 and datediff("d",rst("�ظ�ʱ��"),nowtime)>=1 then 
		msg=msg&rst("����")&"</a></td><td align=center>"
	elseif len(rst("����"))>=15 and datediff("d",rst("�ظ�ʱ��"),nowtime)>=1 then 	
		msg=msg&left(rst("����"),13)&"...</a></td><td align=center>"
	else 
		msg=msg&left(rst("����"),13)&"...</a><img src='../images/new.gif'></td><td align=center>"
	end if
	msg=msg&"<a href='showbbs.asp?bid="&bid&"&fid="&rst("ID")&"' target='_blank' onmouseover="&chr(34)&"window.status='���´��ڲ鿴��������ػظ���';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&"><img src='../images/newwin.gif' border=0></a></td><td>"&showun(rst("����"))&"</td><td>"
	if rst("�ظ�����")="--" then
		msg=msg&len(rst("����"))&"</td><td>"
	else
		msg=msg&"<font color=FF0000>["&rst("�ظ�")&"]</td><td>"
	end if
	if rst("�ظ�����")="--" then
		msg=msg&"--</td><td>"
	else
		msg=msg&showun(rst("�ظ�����"))&"</td><td>"
	end if		
	msg=msg&rst("�ظ�ʱ��")&"</td><td align=right>"&rst("���")&"</td></tr>"
	rst.MoveNext
next
end if
msg=msg&"</table><form method=post action='showessence.asp?bid="&bid&"'><table width=100% bgcolor=cccccc><tr align=center>"
if activepage>1 then
	msg=msg&"<td><a href='showessence.asp?bid="&bid&"&activepage=1' onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='showessence.asp?bid="&bid&"&activepage="&activepage-1&"' onmouseover="&chr(34)&"window.status='ǰһҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">ǰһҳ</a></td>"
else
	msg=msg&"<td>��һҳ</td><td>ǰһҳ</td>"
end if
if activepage<rst.PageCount then
	msg=msg&"<td><a href='showessence.asp?bid="&bid&"&activepage="&activepage+1&"' onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='showessence.asp?bid="&bid&"&activepage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='���һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">���һҳ</a></td>"
else
	msg=msg&"<td>��һҳ</td><td>���һҳ</td>"
end if
msg=msg&"<td>��<input type=text name=activepage size=4 value='"&activepage&"'>ҳ/��"&rst.PageCount&"ҳ<input type=submit value='GO' id=submit1 name=submit1></td></tr></table></form>"&Ba_copyright&"</body></html>"
rst.Close
set rst=nothing
conn.close
set conn=nothing
Response.Write msg
%>
