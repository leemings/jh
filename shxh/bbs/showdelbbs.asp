<%Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1%>
<!--#include file="set.asp"-->
<%
bid=Request.QueryString("bid")
fid=Request.QueryString("fid")
username=session("Ba_jxqy_username")
if instr(Ba_hostname,";"&username&";")=0 then Response.Redirect "../error.asp?id=233"
nowtime=now()
if not (isnumeric(bid) and isnumeric(fid)) then Response.Redirect "../error.asp?id=220"
activepage=Request("activepage")
if activepage="" or not isnumeric(activepage) then activepage=1
activepage=clng(activepage)
msg="<html><head><meta http-equiv='Content-Type' content='text/html; charstet=gb2312'><title>"&Ba_pagename&"</title><link rel='stylesheet' href='../chatroom/css.css'></head><body bgcolor="&Ba_bgcolor&" background="&Ba_bgimage&" oncontextmenu='self.event.returnValue=false' topmargin=0 leftmargin=0 >"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from ��� where id="&bid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=220"
msg=msg&"<table border=0 width=100% bgcolor=cccccc><tr><td>��̳���ƣ�"&showcname(rst("id"),rst("����"),rst("˵��"))&"</td><td>����:"
cname=rst("����")
hostlist=";"&rst("����")&";"
host=split(rst("����"),";")
for i=0 to ubound(host)
	msg=msg&showun(host(i))
next
msg=msg&"</td><td width='30%' ><marquee>"&rst("˵��")&"</marquee></td><td align=right>"&showhomepage(Ba_homepage,Ba_pagename)&"</td></tr></table><p><p>"
rst.Close
rst.Open "select tu.����,tu.�ȼ�,tu.ǩ����,tu.��������,tb.id,tb.��̳,tb.����,tb.����,tb.����,tb.���,tb.�ظ�,tb.����,tb.д��ʱ��,tb.�ظ�����,tb.�ظ�ʱ��,tb.ip,tb.�༭,tb.����,tb.ɾ�� from bbs tb inner join �û� tu on tu.����=tb.���� where tb.ɾ��=true and tb.id="&fid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=200"
if isnull(rst("�ȼ�")) then 
	msg=msg&"���������޷���ѯ�����ݳ���<br>"
else
	author=rst("����")
	content=rst("����")
	sign=rst("ǩ����")
	msg=msg&"<table border=0 width=95% bgcolor=cccccc align=center><tr><td>���±��⣺"&rst("����")&"</td><td align=right>"&autoquote(bid,rst("id"),rst("����"),content)&autosearch(bid,author)&autopro(author)&autouseremail(rst("��������"),author)&autopm(author)&automodify(bid,rst("ID"),username,author)&autoreclaim(bid,rst("id"),author,hostlist,username,Ba_hostname)&autoip(rst("ip"),author)&autoessence(bid,rst("id"),hostlist,username,Ba_hostname)&autodelete(bid,rst("id"),author,hostlist,username,Ba_hostname)&"</td></tr><tr><td colspan=2><table width=100% bgcolor=FFFFFF><tr valign=top><td width=100 align=center>"&author&"<br>"&rst("����")&"<br>"&autolevel(rst("�ȼ�"))&"</td><td>"&Autolink(content)&autoedit(rst("�༭"),rst("����"),rst("�ظ�ʱ��"))&"<br>------------------------<br>"&autolink(sign)&"</td></tr></table></td></tr><tr><td >����:"&autoemail(rst("��������"),author)&"</td><td align=right>ʱ�䣺"&rst("д��ʱ��")&"</td></tr></table><table width=100% bgcolor='"&Ba_bgcolor&"' background='"&Ba_bgimage&"'><tr align=center><td><input type=button onclick=javascript:location.replace('reply.asp?bid="&bid&"&fid="&rst("id")&"'); value=' �� �� ' ); id=button1 name=button1></td><td><input type=button onclick="&chr(34)&"javascript:location.replace('showboard.asp?bid="&bid&"');"&chr(34)&" value=' �� �� '  id=button1 name=button1></td></tr></table>"
end if	
rst.Close
rst.Open "select tu.����,tu.�ȼ�,tu.ǩ����,tu.��������,tb.id,tb.��̳,tb.����,tb.����,tb.����,tb.���,tb.�ظ�,tb.����,tb.д��ʱ��,tb.�ظ�����,tb.�ظ�ʱ��,tb.ip,tb.�༭,tb.����,tb.ɾ�� from bbs tb inner join �û� tu on tu.����=tb.���� where tb.ɾ��=true and tb.����="&fid&" order by tb.д��ʱ��",conn
do while not (rst.EOF or rst.BOF)
	if isnull(rst("�ȼ�")) then 
		msg=msg&"���������޷���ѯ�����ݳ���<br>"
	else
		author=rst("����")
		content=rst("����")
		sign=rst("ǩ����")
		msg=msg&"<table border=0 width=95% bgcolor=cccccc align=center><tr><td>���±��⣺"&rst("����")&"</td><td align=right>"&autoquote(bid,rst("id"),rst("����"),content)&autosearch(bid,author)&autopro(author)&autouseremail(rst("��������"),author)&autopm(author)&automodify(bid,rst("ID"),username,author)&autoreclaim(bid,rst("id"),author,hostlist,username,Ba_hostname)&autoip(rst("ip"),author)&autodelete(bid,rst("id"),author,hostlist,username,Ba_hostname)&"</td></tr><tr><td colspan=2><table width=100% bgcolor=FFFFFF><tr valign=top><td width=100 align=center>"&author&"<br>"&rst("����")&"<br>"&autolevel(rst("�ȼ�"))&"</td><td>"&Autolink(content)&autoedit(rst("�༭"),rst("����"),rst("�ظ�ʱ��"))&"<br>------------------------<br>"&autolink(sign)&"</td></tr></table></td></tr><tr><td >����:"&autoemail(rst("��������"),author)&"</td><td align=right>ʱ�䣺"&rst("д��ʱ��")&"</td></tr></table><br>"
	end if	
	rst.MoveNext 
loop
rst.Close
set rst=nothing
conn.execute "update bbs set ���=���+1 where id="&fid
conn.close
set conn=nothing
msg=msg&Ba_copyright&"</body></html>"
Response.Write msg
%>
