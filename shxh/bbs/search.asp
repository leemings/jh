<%Response.Expires=-1%>
<!--#include file="set.asp"-->
<%
searchstr=Request("searchstr")
bid=Request("bid")
activepage=Request("activepage")
if activepage="" or not isnumeric(activepage) then activepage=1
activepage=clng(activepage)
if activepage<1 then activepage=1
msg="<html><head><meta http-equiv='Content-Type' content='text/html; charstet=gb2312'><title>"&Ba_pagename&"</title><link rel='stylesheet' href='../chatroom/css.css'></head><body bgcolor="&Ba_bgcolor&" background="&Ba_bgimage&" oncontextmenu='self.event.returnValue=false' topmargin=0 leftmargin=0 >"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open "select * from ��� where id="&bid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=220"
msg=msg&"<table border=0 width=100% bgcolor=cccccc><tr><td>��̳���ƣ�"&showcname(rst("id"),rst("����"),rst("˵��"))&"</td><td>����:"
host=split(rst("����"),";")
for i=0 to ubound(host)
	msg=msg&showun(host(i))
next
msg=msg&"</td><td width='30%' ><marquee>"&rst("˵��")&"</marquee></td><td align=right>"&showhomepage(Ba_homepage,Ba_pagename)&"</td></tr></table><p><table width=100% bgcolor=cccccc border=1 align=center><tr bgcolor=fffe48 align=center><td>&nbsp;</td><td>&nbsp;</td><td>���±���</td><td>�´���</td><td>����</td><td>��[��]</td><td>���ظ�</td><td>������ʱ��</td><td>���</td></tr>"
rst.Close
rst.Open "select * from bbs where (���� like '%"&searchstr&"%' or ���� like '%"&searchstr&"%' or ���� like '%"&searchstr&"%') and ɾ��=false and ��̳="&bid&" order by �ظ�ʱ��  desc",conn,1,3
if  rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=226"
bgc="bgcolor=f7f7f7"
searchlist=" "
pagenum=Ba_pagesize
i=0
do while not rst.EOF 
	if rst("����")=0  and instr(searchlist," "&rst("id")&" ")=0 and i>=(activepage-1)*pagenum and i<activepage*pagenum then
	i=i+1
	searchlist=searchlist&rst("id")&" "
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
		msg=msg&"<img src=../images/essence.gif></td><td><a href='showbbs.asp?bid="&bid&"&fid="&rst("ID")&"' onmouseover="&chr(34)&"window.status='�鿴��������ػظ���';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"
	else
		msg=msg&"&nbsp;</td><td><a href='showbbs.asp?bid="&bid&"&fid="&rst("ID")&"'  onmouseover="&chr(34)&"window.status='�鿴��������ػظ���';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"
	end if
	if len(rst("����"))<15 and datediff("h",rst("�ظ�ʱ��"),nowtime)<1 then 
		msg=msg&rst("����")&"</a><img src='../images/new.gif'></td><td>"
	elseif len(rst("����"))<15 and datediff("h",rst("�ظ�ʱ��"),nowtime)>=1 then 
		msg=msg&rst("����")&"</a></td><td>"
	elseif len(rst("����"))>=15 and datediff("h",rst("�ظ�ʱ��"),nowtime)>=1 then 	
		msg=msg&left(rst("����"),13)&"...</a></td><td>"
	else 
		msg=msg&left(rst("����"),13)&"...</a><img src='../images/new.gif'></td><td >"
	end if
	msg=msg&"<a href='showbbs.asp?bid="&bid&"&fid="&rst("ID")&"' target='_blank' onmouseover="&chr(34)&"window.status='���´��ڲ鿴��������ػظ���';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&"><img src='../images/newwin.gif'></a></td><td>"&showun(rst("����"))&"</td><td>"
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
	elseif rst("����")<>0 and instr(searchlist," "&rst("����")&" ")=0  and i>=(activepage-1)*pagenum and i<activepage*pagenum then
	i=i+1	
	searchlist=searchlist&rst("����")&" "
	if bgc="bgcolor=f7f7f7" then 
		bgc=""
	else
		bgc="bgcolor=f7f7f7"
	end if
	set rst2=server.CreateObject("adodb.recordset")
	rst2.Open "select * from bbs where id="&rst("����"),conn
	msg=msg&"<tr "&bgc&"><td>"
	if datediff("d",rst2("�ظ�ʱ��"),nowtime)>=1 and rst2("�ظ�")<Ba_hotbbsnum then
		msg=msg&"<img src='../images/closed.gif'></td><td>"
	elseif	datediff("d",rst2("�ظ�ʱ��"),nowtime)>=1 and rst2("�ظ�")>=Ba_hotbbsnum then
		msg=msg&"<img src='../images/hotclosed.gif'></td><td>"
	elseif	datediff("d",rst2("�ظ�ʱ��"),nowtime)<1 and rst2("�ظ�")<Ba_hotbbsnum then
		msg=msg&"<img src='../images/closedb.gif'></td><td>"
	else
		msg=msg&"<img src='../images/hotclosedb.gif'></td><td>"
	end if	
	if rst2("����")=True then
		msg=msg&"<img src=images/essence.gif></td><td><a href='showbbs.asp?bid="&bid&"&fid="&rst2("ID")&"' onmouseover="&chr(34)&"window.status='�鿴��������ػظ���';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"
	else
		msg=msg&"&nbsp;</td><td><a href='showbbs.asp?bid="&bid&"&fid="&rst2("ID")&"'  onmouseover="&chr(34)&"window.status='�鿴��������ػظ���';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"
	end if
	if len(rst2("����"))<15 and datediff("h",rst2("�ظ�ʱ��"),nowtime)<1 then 
		msg=msg&rst2("����")&"</a><img src='../images/new.gif'></td><td>"
	elseif len(rst2("����"))<15 and datediff("h",rst2("�ظ�ʱ��"),nowtime)>=1 then 
		msg=msg&rst2("����")&"</a></td><td>"
	elseif len(rst2("����"))>=15 and datediff("h",rst2("�ظ�ʱ��"),nowtime)>=1 then 	
		msg=msg&left(rst2("����"),13)&"...</a></td><td>"
	else 
		msg=msg&left(rst2("����"),13)&"...</a><img src='../images/new.gif'></td><td >"
	end if
	msg=msg&"<a href='showbbs.asp?bid="&bid&"&fid="&rst2("ID")&"' target='_blank' onmouseover="&chr(34)&"window.status='���´��ڲ鿴��������ػظ���';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&"><img src='../images/newwin.gif'></a></td><td>"&showun(rst2("����"))&"</td><td>"
	if rst2("�ظ�����")="--" then
		msg=msg&len(rst2("����"))&"</td><td>"
	else
		msg=msg&"<font color=FF0000>["&rst2("�ظ�")&"]</td><td>"
	end if
	if rst2("�ظ�����")="--" then
		msg=msg&"--</td><td>"
	else
		msg=msg&showun(rst2("�ظ�����"))&"</td><td>"
	end if		
	msg=msg&rst2("�ظ�ʱ��")&"</td><td align=right>"&rst2("���")&"</td></tr>"
	rst2.Close
	set rst2=nothing
	elseif rst("����")=0  and instr(searchlist," "&rst("id")&" ")=0 and (i<(activepage-1)*pagenum or i>=activepage*pagenum) then
	i=i+1
	searchlist=searchlist&rst("id")&" "
	elseif rst("����")<>0 and instr(searchlist," "&rst("����")&" ")=0  and (i<(activepage-1)*pagenum or i>=activepage*pagenum) then
	i=i+1
	searchlist=searchlist&rst("����")&" "
	end if
	rst.MoveNext
loop
if i mod pagenum =0 then
	pagecount=i/pagenum
else
	pagecount=i\pagenum+1
end if	
msg=msg&"</table><form method=post action='search.asp?bid="&bid&"&searchstr="&searchstr&"' id=form1 name=form1><table width=100% bgcolor=cccccc>"
if activepage>1 then
	msg=msg&"<td><a href='search.asp?bid="&bid&"&searchstr="&searchstr&"&activepage=1' onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='search.asp?bid="&bid&"&searchstr="&searchstr&"&activepage="&activepage-1&"' onmouseover="&chr(34)&"window.status='ǰһҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">ǰһҳ</a></td>"
else
	msg=msg&"<td>��һҳ</td><td>ǰһҳ</td>"
end if
if activepage<pagecount then
	msg=msg&"<td><a href='search.asp?bid="&bid&"&searchstr="&searchstr&"&activepage="&activepage+1&"' onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='search.asp?bid="&bid&"&searchstr="&searchstr&"&activepage="&pagecount&"' onmouseover="&chr(34)&"window.status='���һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">���һҳ</a></td>"
else
	msg=msg&"<td>��һҳ</td><td>���һҳ</td>"
end if
msg=msg&"<td>��<input type=text name=activepage size=4 value='"&activepage&"'>ҳ/��"&pagecount&"ҳ<input type=submit value='GO' id=submit1 name=submit1></td></tr></table></form>"&Ba_copyright&"</body></html>"

rst.Close
set rst=nothing
conn.close
set conn=nothing

Response.Write msg
%>