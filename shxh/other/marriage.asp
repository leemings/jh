<%
Response.Expires=-1
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016" 
sex=session("Ba_jxqy_usersex")
if sex="��" then 
	sex="��"
else
	sex="��"
end if		
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
opt=Request.QueryString("opt")
if opt="" then opt="���"
acpage=request("acpage")
if not isnumeric(acpage) then acpage=1
acpage=clng(acpage)
if acpage<1 then acpage=1
msg="<head><title>ý��</title><link rel=stylesheet href='../chatroom/css.css'></head><body  bgcolor='"&bgcolor&"' background='"&bgimage&"' oncontextmenu='self.event.returnValue=false;'><table width=40% border=3 align=center bgcolor=FFB442><tr align=center>"
if opt="���" then
	msg=msg&"<td bgcolor=FFFF00><a href='marriage.asp?opt=���' onmouseover="&chr(34)&"window.status='�����������̹¶�����ɣ�����ף���㣡';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title='�����������̹¶�����ɣ�����ף���㣡'>���</a></td><td><a href='marriage.asp?opt=���' onmouseover="&chr(34)&"window.status='����ԭ����ģ��ҿ����ĺ��ӣ�';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title='����ԭ����ģ��ҿ����ĺ��ӣ�'>���</a></td></tr></table><p align=center><b><font color=CC0000 face=��Բ size=4>ý �� ��</font></b></p><table width=100% border=3><tr bgcolor=FFFF00 align=center><td>�����</td><td>������</td><td>������</td><td>����</td></td></tr>"
	sqlstr="select  * from ý�� where ����=False order by ʱ�� desc"
else
	msg=msg&"<td><a href='marriage.asp?opt=���' onmouseover="&chr(34)&"window.status='�����������̹¶�����ɣ�����ף���㣡';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title='�����������̹¶�����ɣ�����ף���㣡'>���</a></td><td bgcolor=FFFF00><a href='marriage.asp?opt=���' onmouseover="&chr(34)&"window.status='����ԭ����ģ��ҿ����ĺ��ӣ�';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title='����ԭ����ģ��ҿ����ĺ��ӣ�'>���</a></td></tr></table><p align=center><b><font color=CC0000 face=��Բ size=4>ý �� ��</font></b></p><table width=100% border=3><tr bgcolor=FFFF00 align=center><td>������</td><td>������</td><td>�������</td><td>����</td></tr>"
	sqlstr="select  * from ý�� where ����=True order by ʱ�� desc"
end if
set conn=server.createobject("adodb.connection") 
conn.open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.open sqlstr,conn,1,2
if not (rst.EOF or rst.BOF) then
	rst.Pagesize=10
	if acpage>rst.PageCount then acpage=rst.PageCount
	rst.AbsolutePage=acpage
	for i=1 to rst.PageSize
		if rst.EOF then exit for
		msg=msg&"<tr><td>"&rst("������")&"</td><td>"&rst("����")&"</td><td>"&rst("˵��")&"</td><td>"
		if rst("����")=username then 
			msg=msg&"<a href='modmar.asp?id="&rst("id")&"' onmouseover="&chr(34)&"window.status='"&sex&"˵�ĺ���ܶԣ��һ���ͬ���˰�';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title='"&sex&"˵�ĺ���ܶԣ��һ���ͬ���˰�'>��Ӧ</a>"
		elseif rst("������")=username then 
			msg=msg&"<a href='delmar.asp?id="&rst("id")&"' onmouseover="&chr(34)&"window.status='���뻹�����˰�';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title='���뻹�����˰�'>ɾ��</a>"
		else
			msg=msg&"����"
		end if
		msg=msg&"</td></tr>"
		rst.movenext
	next
end if
msg=msg+"</table><form action='marriage.asp?opt="&opt&"' method=post id=form1 name=form1><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=ffff00><tr align=center>"
if acpage>1 then
	msg=msg&"<td><a href='marriage.asp?opt="&opt&"&acpage=1' onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='marriage.asp?opt="&opt&"&acpage="&acpage-1&"' onmouseover="&chr(34)&"window.status='ǰһҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">ǰһҳ</a></td>"
else
	msg=msg&"<td>��һҳ</td><td>ǰһҳ</td>"
end if
if acpage<rst.PageCount then
	msg=msg&"<td><a href='marriage.asp?opt="&opt&"&acpage="&acpage+1&"' onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='marriage.asp?opt="&opt&"&acpage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='���һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">���һҳ</a></td>"
else
	msg=msg&"<td>��һҳ</td><td>���һҳ</td>"
end if
msg=msg&"<td>��<input type=text name=acpage size=4 value='"&acpage&"'>ҳ��"&rst.PageCount&"ҳ<input type=submit value='GO' id=submit1 name=submit1></td></tr></table></form>"
rst.Close
rst.Open "select ��ż from �û� where ����='"&username&"'",conn
mate=rst("��ż")
if opt="���" and mate="��" then
	msg=msg&"<form action='addmar.asp' method=post><table align=center><tr><td>������<input type=text name=quest size=10 maxlength=7>������<input type=text name=content size=20 maxlength=50><input type=submit name=opt value='���'></td></tr></table></form>"
elseif opt="���" and mate<>"��" then
	msg=msg&"<form action='addmar.asp' method=post><table align=center><tr><td>��ż��<input type=text name=quest size=10 maxlength=7 readonly value="&mate&">�ҵ�ԭ��<input type=text name=content size=20 maxlength=50><input name=opt type=submit value='���'></td></tr></table></form>"
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
msg=msg&"<p align=right><a href='#' onclick='javascript:window.close();'>�ر�</a></p></body>"
Response.Write msg
%>