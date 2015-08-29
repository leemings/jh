<!-- #include file="shopconn.asp" -->
<!-- #include file="int.asp" -->
<html>
<body>
<%username=trim(request("username"))
if request("sex")="ÄÐ" then mysex=1 else mysex=0 end if
set rs=server.createobject("adodb.recordset")
sql="select * from [User] where username='"&username&"'"
rs.open sql,connshop,1,1	
if rs.eof then
if sex=1 then%>
<br><table border=0 cellspacing=0 cellpadding=0  align="center"><tr><td><img src="pic/nan.gif"></td></table>
<%else%>
<br><table border=0 cellspacing=0 cellpadding=0  align="center"><tr><td><img src="pic/nv.gif"></td></talbe>
<%end if
else
uservisual=rs("visual")%>
<br><table border=0 cellspacing=0 cellpadding=0  align="center"><tr><td align=center visualwidth=111 visualheight=180 ImagePath='<%=PicPath%>' usergender=<%=mysex%> baseName='/jhface/userface/left-bg.gif' visual='<%=uservisual%>' style=behavior:url(/jhface/usershow.htc) localpic='<%=localpic%>' userface='/jhface/userface/stop.gif' facewidth=74 faceheight=180></td></tr></table>
<%end if
rs.close
set rs=nothing
connshop.close
set connshop=nothing
%>
</body>
</html>