<!-- #include file="shopconn.asp" -->
<!--#include file="inc.asp"-->
<%dim username,uservisual
username=session("sjjh_name")
dim rs
set rs=server.createobject("adodb.recordset")
sql="select visual from [User] where username='"&UserName&"'"
rs.open sql,connshop,1,1
if rs.eof and rs.bof then%>
<p align="center"><a href="show.asp" target="_blank">设置我的虚拟人物<%=username%></a></p>
<%else
uservisual=rs("Visual")%>

<table border=0 cellspacing=0 cellpadding=0 width=111><tr><td align=center visualwidth=111 visualheight=180 ImagePath="<%=PicPath%>" usergender=<%=mysex%> baseName="userface/left-bg.gif" visual="<%=uservisual%>" style=behavior:url(usershow.htc) localpic="<%=localpic%>" userface="userface/stop.gif" facewidth=74 faceheight=180></td></tr></table>

<%end if
rs.close
set rs=nothing
connshop.close
set connshop=nothing%>