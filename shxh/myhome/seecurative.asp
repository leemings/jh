<%
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select ����,����,����,��Ч,���� from ��Ʒ where ����='ҩƷ' and ������='"&username&"' and ����>0 order by �۸�"
rst.Open sqlstr,conn
msg="<table border='1' cellspacing='0' cellpadding='1' bordercolor='#800000' bordercolorlight='#800000' bordercolordark='#FFFFFF' width='100%'><tr align=center bgcolor=ffff00><td>����</td><td>����</td><td>����</td><td>��Ч</td><td>����</td></tr>"
do while not (rst.BOF or rst.EOF)
	cname=rst("����")
	msg=msg&"<tr><td ><a href='usecur.asp?mg="&cname&"'>"&cname&"</td><td>"&rst("����")&"</td><td>"&rst("����")&"</td><td>"&rst("��Ч")&"</td><td>"&rst("����")&"</td></tr>"
	rst.MoveNext
loop
msg=msg&"</table>"
rst.Close 
set rst=nothing
conn.Close
set conn=nothing
%>
<head><link rel="stylesheet" href="../chatroom/css.css"></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false marginwidth="5" marginheight="0">
<div align=center>
<font color=0000FF size=4>ӵ��ҩƷ</font><hr>
<%=msg%>
<input type=button value='����' onclick=javascript:location.href='myhome.asp';>
</div>
</body>
