<%
Response.Expires=-1
chatroomsn=Session("Ba_jxqy_userchatroomsn")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
allonline=split(Application("Ba_jxqy_allonline"),";")
allonlineubd=ubound(allonline)
%>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<div align=center><font size=5 color=0000ff><%=Application("Ba_jxqy_systemname"&chatroomsn)%></font><br><font color=FF0000><%=Application("Ba_jxqy_onlinenum"&chatroomsn)%></font>/<font color=008800><%=Application("Ba_jxqy_allonlinenum")%></font>人在线<hr>
<form action='sendwebicq.asp' method=post>
<table width='100%' align=center>
<tr align=center><td><font color=0000FF>飞鸽传书</font></td></tr>
<tr align=center><td><select name=sendto>
<%for i=1 to allonlineubd-1 
Response.Write "<option value='"&allonline(i)&"'>"&allonline(i)&"</option>"
next
%></select></td></tr>
<tr><td>
<textarea name=message rows=15 cols=16>值此良晨美景，谨此送上我最深的祝福：</textarea>
</td></tr>
<tr align=center><td><input type=submit value='放出信鸽'></td></tr>
</table>
</table>
</form>
</div>
</body>