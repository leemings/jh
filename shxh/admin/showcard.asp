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
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from cardshop order by cprice",conn
do while not rst.EOF
	cname=trim(rst("cname"))
	msg=msg&"<tr><td>"&cname&"</td><td>"&rst("cespecial")&"</td><td>"&rst("ctime")&"</td><td align=right>"&rst("cprice")&"</td><td><a href='upcard1.asp?opt=�޸�&cid="&rst("id")&"' onmouseover="&chr(34)&"window.status='�޸�"&cname&"���й�����';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">�޸�</a> | <a href='upcard1.asp?opt=ɾ��&cid="&rst("id")&"' onmouseover="&chr(34)&"window.status='ɾ��"&cname&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">ɾ��</a></td></tr>"
	rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<html>
<head>
<link rel=stylesheet href='../style.css'>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor='<%=bgcolor%>' background='<%=bgimage%>'>
<p align=center><font color=0000ff size=4>���߹���</font></p>����Ա���ܴ�ʱֻ�л�Ա���Թ����ʹ�õ��ߣ���֮ȫ���û�����ʹ��<hr>
<table width=80% border=3 align=center>
<tr bgcolor=ffff00 align=center><td>��������</td><td>����</td><td>����ʱ��(S)</td><td>�۸�</td><td><a href='upcard1.asp?opt=����&cid=-1' onmouseover="window.status='������Ƭ����';return true;" onmouseout="window.status='';return true;">����</a></td></tr>
<%=msg%>
</table>
</body>
</html>