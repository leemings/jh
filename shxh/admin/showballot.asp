<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("Ba_jxqy_username")
usergrade=session("Ba_jxqy_usergrade")
usercorp=session("BA_jxqy_usercorp")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
allright=Application("Ba_jxqy_allright")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "ballotsystem",conn
do while not rst.EOF
	id=rst("id")
	condition=rst("condition")
	active=rst("active")
	if condition="true" then
		condition="������"
	elseif condition="false" then
		condition="��ֹ�κ���ͶƱ"
	else		
		condition=replace(condition,">=","��С��")
		condition=replace(condition,"<=","������")
		condition=replace(condition,">","����")
		condition=replace(condition,"<","����")
		condition=replace(condition,"=","Ϊ")
		condition=replace(condition," and ","����")
		condition=replace(condition," or ","��")
	end if
	explain=replace(rst("explain"),chr(13)&chr(10),"<br>")
	if allright=usergrade then
		msg=msg&"<tr><td>"&id&"</td><td>"&rst("title")&"</td><td>"&explain&"</td><td>"&rst("deadline")&"</td><td>"&condition&"</td><td>"&active&"</td><td><a href=chgballot.asp?id="&id&"&opt=�޸�>�޸�</a> | <a href=chgballot.asp?id="&id&"&opt=ɾ��>ɾ��</a> | <a href='showblt.asp?pid="&id&"'>�鿴</a></td></tr>"
	else
		msg=msg&"<tr><td>"&id&"</td><td>"&rst("title")&"</td><td>"&explain&"</td><td>"&rst("deadline")&"</td><td>"&condition&"</td><td>"&active&"</td></tr>"
	end if	
	rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head>
<link rel='stylesheet' href='../chatroom/css.css'>
</head>
<body oncontextmenu=self.event.returnValue=false topmargin=0 background='<%=bgimage%>' bgcolor='<%=bgcolor%>'>
<p align=center>ͶƱϵͳ</p>
<hr>
<table border=3 width=80% align=center>
<tr bgcolor=#FFFF00 align=center><td>ID</td><td>����</td><td>˵��</td><td>��ֹ����</td><td>ͶƱ����</td><td>��Ծ</td><%if usergrade=allright and usercorp="�ٸ�" then%><td><a href=chgballot.asp?opt=����&id=-1>����ͶƱ��</a></td><%end if%></tr>
<%=msg%>
</table>
</body>