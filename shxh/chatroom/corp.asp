<%
chatroomsn=Session("Ba_jxqy_userchatroomsn")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select ����,��� from ���� where ��������<>'False' order by ����ϵ��"
rst.Open sqlstr,conn
do while not (rst.EOF or rst.BOF)
	corp=rst("����")
	msg=msg+"<tr><td bgcolor=ffff00 align=center><a href='joincorp.asp?mg="&corp&"' omouseover="&chr(34)&"window.status='����"&corp&"';return true;"&chr(34)&"  onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title="&rst("���")&">��"&corp&"��</a></td></tr>"
rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head><link rel="stylesheet" href="css.css">
</head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false leftmargin="5" marginwidth="5">
<div align=center><font size=5 color=0000ff><%=Application("Ba_jxqy_systemname"&chatroomsn)%></font><br><font color=FF0000><%=Application("Ba_jxqy_onlinenum"&chatroomsn)%></font>/<font color=008800><%=Application("Ba_jxqy_allonlinenum")%></font>������
  <hr noshade size="1" color=red>
  <font color=0000FF><b>��������</b></font><br>
  <font color="#FF0000"><b>[</b>ͣ���ɿ�˵����������<b>]</b></font><br>
 
    <table width=98% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF>
<%=msg%>
</table>

</div>
</body>
