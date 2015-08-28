<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
chatroomsn=Session("yx8_mhjh_userchatroomsn")
allonline=split(Application("yx8_mhjh_allonline"),";")
allonlineubd=ubound(allonline)
%>
<head>
<link rel='stylesheet' href='css2.css'>
</head>
<body  oncontextmenu=self.event.returnValue=false background='bg1.gif'>
<div align=center><font color=008800><%=Application("yx8_mhjh_allonlinenum")%></font>人在线
<form action='sendwebicq.asp' method=post>
<table width='100%' align=center>
<tr align=center><td><font color=0000FF>飞鸽传书</font></td></tr>
<tr align=center><td><select name=sendto>
<%for i=1 to allonlineubd-1
if allonline(i)<>Application("yx8_mhjh_admin") then
Response.Write "<option value='"&allonline(i)&"'>"&allonline(i)&"</option>"
end if
next
%></select></td></tr>
<tr><td>
<textarea name=message rows=15 cols=15></textarea>
</td></tr>
<tr align=center><td><input type=submit value='放出信鸽'></td></tr>
</table>
</form>
</div>
</body>
