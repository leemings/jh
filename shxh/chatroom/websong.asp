<%
Response.Expires=-1
chatroomsn=Session("Ba_jxqy_userchatroomsn")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
allonline=split(Application("Ba_jxqy_allonline"),";")
allonlineubd=ubound(allonline)
dim sng()
dim fnm()
i=0
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "mid",conn
do while not rst.EOF
	redim preserve sng(i),fnm(i)
	sng(i)=Trim(rst("songname"))
	fnm(i)=Trim(rst("filename"))
	i=i+1
	rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close 
set conn=nothing
%>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<div align=center><font size=5 color=0000ff><%=Application("Ba_jxqy_systemname"&chatroomsn)%></font><br><font color=FF0000><%=Application("Ba_jxqy_onlinenum"&chatroomsn)%></font>/<font color=008800><%=Application("Ba_jxqy_allonlinenum")%></font>人在线<hr>
<form action='sendwebsng.asp' method=post>
<table width='100%' align=center>
<tr align=center><td><font color=0000FF>好歌送好友</font></td></tr>
<tr align=center><td><select name=sendto>
<%for i=1 to allonlineubd-1 
Response.Write "<option value='"&allonline(i)&"'>"&allonline(i)&"</option>"
next
%></select></td></tr>
<tr><td align=center><select name='song'>
<%for i=0 to ubound(sng)%>
<option value="歌曲<%=sng(i)%><bgsound src='../mid/song/<%=fnm(i)%>' loop=1>"><%=sng(i)%></option>
<%next%>
</select></td></tr>
<tr><td>
<textarea name=message rows=15 cols=16>好歌送好友！</textarea>
</td></tr>
<tr align=center><td><input type=submit value='点歌'></td></tr>
</table>
</table>
</form>
</div>
</body>