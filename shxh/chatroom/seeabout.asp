<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
chatroomsn=Session("Ba_jxqy_userchatroomsn")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from �û� where ����='"&username&"'",conn
for i=11 to 25
	msg=msg&"<tr><td><font color=FF0000>"&rst.Fields(i).Name&"</font></td><td>"&rst.Fields(i).Value&"</td></tr>"
next
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false leftmargin="5" marginwidth="5">
<LINK href="css.css" rel=stylesheet>
<div align=center><font size="4" color="#CC0000" face="��Բ"><b><%=Application("Ba_jxqy_systemname"&chatroomsn)%></b></font><br>
  <font color=FF0000><%=Application("Ba_jxqy_onlinenum"&chatroomsn)%></font>/<font color=008800><%=Application("Ba_jxqy_allonlinenum")%></font>������<hr>
<font color=0000FF>����״̬</font><br>
  <table width=95% bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF>
    <tr>
      <td height="19" valign="top"><%=msg%></td>
    </tr>
  </table>
</div>
</body>
