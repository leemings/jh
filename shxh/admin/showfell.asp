<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("Ba_jxqy_username")
usergrade=session("Ba_jxqy_usergrade")
usercorp=session("Ba_jxqy_usercorp")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usercorp="�ٸ�" and usergrade=Application("Ba_jxqy_allright")) then Response.Redirect "../error.asp?id=500"
nowtime=now()
nowtimetype="#"&month(nowtime)&"/"&day(nowtime)&"/"&year(nowtime)&"#"
activepage=Request("activepage")
if activepage="" or not isnumeric(activepage) then activepage=1
activepage=clng(activepage)
if activepage<1 then activepage=1
set conn=server.CreateObject("adodb.connection")
conn.open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select ����,����,���,�ȼ�,��Աʱ�� from �û� where ��Ա=true and ��Աʱ��>"&nowtimetype&" order by ��Աʱ�� desc",conn,1,2
if rst.EOF then
	msg="<tr align=center><td>����û�����ǻ�Ա</td></tr>"
else
	rst.PageSize=20
	if activepage>rst.PageCount then activepage=rst.PageCount
	rst.AbsolutePage=activepage
	for i=1 to rst.Pagesize
	if rst.EOF or rst.BOF then exit for
		msg=msg&"<tr><td><a href='#' onclick=javacript:document.form2.username.value='"&rst("����")&"'; onmouserover="&chr(34)&"window.status='ѡ���������Ϊ"&rst("����")&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"&rst("����")&"</a></td><td>"&rst("����")&"</td><td>"&rst("���")&"</td><td>"&rst("�ȼ�")&"</td><td>"&formatdatetime(rst("��Աʱ��"),1)&"</td></tr>"
		rst.MoveNext
	next
	msg=msg&"</table><table border=3 width='100%'><form action='showfell.asp' method=post id=form1 name=form1><tr align=center bgcolor=cccccc>"
		if activepage>1 then
	msg=msg&"<td><a href='showfell.asp?activepage=1' onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='showfell.asp?activepage="&activepage-1&"' onmouseover="&chr(34)&"window.status='ǰһҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">ǰһҳ</a></td>"
	else
		msg=msg&"<td>��һҳ</td><td>ǰһҳ</td>"
	end if
	if activepage<rst.PageCount then
		msg=msg&"<td><a href='showfell.asp?activepage="&activepage+1&"' onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='showfell.asp?activepage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='���һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">���һҳ</a></td>"
	else
		msg=msg&"<td>��һҳ</td><td>���һҳ</td>"
	end if
	msg=msg&"<td>��<input type=text name=activepage size=4 value='"&activepage&"'>ҳ/��"&rst.PageCount&"ҳ<input type=submit value='GO' id=submit1 name=submit1></td></tr></form>"
end if
%>
<html>
<head>
<link rel='stylesheet' href='../style.css'>'
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor='<%=bgcolor%>' background='<%=bgimage%>'>
<p align=center><font color=0000ff size=4>��Ա����</font></p><hr>
<table width=80% align=center border=3><tr bgcolor=ffff00 align=center><td>����</td><td>����</td><td>���</td><td>�ȼ�</td><td>��Ա����ʱ��</td></tr><%=msg%></table>
<form action=upfellow.asp method=post name=form2><table width=80% align=center border=3><tr><td>����</td><td><input type=text name='username' value='' maxlength=7 size=14></td><td>ʱ��</td><td> <input type=text name='month' value='1' maxlength=2 size=2>����</td><td colspan=2> <input type=submit value=' �� �� ' name=submit> <input type=submit value=' �� �� ' name=submit> <input type=reset value=' �� �� '> </td></td></tr></table></form>
</body>
</html>