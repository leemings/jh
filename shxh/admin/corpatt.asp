<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
usercorp=Session("Ba_jxqy_usercorp")
chatroomsn=Session("Ba_jxqy_userchatroomsn")
if username="" then Response.redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from �û� where ����='"&username&"' and ����='"&usercorp&"' and ���='����'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=046"
userenergy=rst("����")
rst.Close
rst.Open "select * from ��ʽ where ����='"&usercorp&"' order by ��������",conn
do while not rst.EOF 
	msg=msg&"<tr><td>"&rst("��ʽ")&"</td><td>"&rst("���ľ���")&"</td><td>"&rst("��������")&"</td><td>"&rst("��������")&"</td><td>"&rst("��Ч")&"</td></tr>"
	rst.MoveNext 
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head>
<link rel='stylesheet' href='../chatroom/style1.css'>
<script language=javascript>var userenergy=<%=userenergy\20%>;function energychg(){if(document.form1.energy.value>userenergy){alert('���ľ������㣡');}else{if(document.form1.especial.value=='��'){document.form1.mp.value=parseInt(document.form1.energy.value/10);}else{document.form1.mp.value=document.form1.energy.value;}document.form1.attack.value=document.form1.energy.value;}}</script>
</head>
<body oncontextmenu=self.event.returnValue=false background='<%=bgimage%>' bgcolor='<%=bgcolor%>' ><p align=center><font color=ff0000 size=5>�Ż�����</font><hr></p><form action=corpatt2.asp method=post id=form1 name=form1><table width=80% align=center border=3><tr align=center bgcolor=00ff00><td>��ʽ</td><td>����</td><td>����</td><td>����</td><td>��Ч</td></tr><%=msg%>
<tr><td><input type=text name=attackname value='��ʽ����' maxlength=7 size=14 ></td><td><input type=text name=energy value='<%=userenergy\20%>' maxlength=5 size=5  onchange='javascript:energychg();'></td><td><input type=text name=mp value='<%=userenergy\200%>' maxlength=5 size=5 readonly></td><td><input type=text name=attack value='<%=userenergy\20%>' maxlength=5 size=5 readonly></td><td><select name=especial onchange='javascript:energychg();'><option value='��' selected>��</option><option value='���'>���</option><option value='�ж�'>�ж�</option><option value='����'>����</option><option value='���'>���</option></select></td></tr></table><p align=center><input type=submit value='�½�' id=submit1 name=submit1></p></form></body>