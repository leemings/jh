<%
Response.Expires=-1
chatroomsn=Session("Ba_jxqy_userchatroomsn")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select ����,��Ч,���� from ��Ʒ where ����='����' and ������='"&username&"' and ����>0 order by �۸�"
rst.Open sqlstr,conn
msg="<table width=95% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=1 bordercolordark=FFFFFF><tr></tr>"
do while not (rst.BOF or rst.EOF)
	msg=msg&"<tr><td><a href=javascript:parent.talkfrm.settalk('//Ͷ��','"&rst("����")&"') target='talkfrm' onmouseover=""window.status='ʹ�ð���';return true;"" onmouseout=""window.status='';return true;"" title="&chr(34)&rst("��Ч")&chr(34)&">"&rst("����")&"</a></td><td width=15  bgcolor=ffff00>"&rst("����")&"</td></tr>"
	rst.MoveNext
loop
msg=msg&"</table>"
rst.Close 
set rst=nothing
conn.Close
set conn=nothing
%>
<head><link rel="stylesheet" href="style1.css"></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false leftmargin="4" marginwidth="4">
<div align=center><font size="4" color="#CC0000" face="��Բ"><b><%=Application("Ba_jxqy_systemname"&chatroomsn)%></b></font><br>
  <font color=FF0000><%=Application("Ba_jxqy_onlinenum"&chatroomsn)%></font>/<font color=008800><%=Application("Ba_jxqy_allonlinenum")%></font>������
  <hr noshade size="1" color=red>
<font color=0000FF>ӵ�а���</font><br>
<%=msg%>
</div>
</body>
