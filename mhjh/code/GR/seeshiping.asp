<%
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select ����,����,����,��Ч,���� from ��Ʒ where ����='ʳƷ' and ������='"&username&"' and ����>0 order by �۸�"
rst.Open sqlstr,conn
msg="<table border='1' cellspacing='0' cellpadding='1' bordercolor='#800000' bordercolorlight='#800000' bordercolordark='#FFFFFF' width='100%'><tr align=center bgcolor=ffffff><td>����</td><td>����</td><td>����</td><td>��Ч</td><td>����</td></tr>"
do while not (rst.BOF or rst.EOF)
cname=rst("����")
msg=msg&"<tr><td ><a href='usecur2.asp?mg="&cname&"'>"&cname&"</td><td>"&rst("����")&"</td><td>"&rst("����")&"</td><td>"&rst("��Ч")&"</td><td>"&rst("����")&"</td></tr>"
rst.MoveNext
loop
msg=msg&"</table>"
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head><link rel="stylesheet" href="../style.css"></head>
<body background='../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false marginwidth="5" marginheight="0">
<div align=center>
<b>
<font color="#000000" size="4" face="����">�ҵ�ʳƷ</font></b><hr>
<%=msg%>
</div>
</body>
