<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_down_conn.asp"-->
<!--#include file="inc/FORUM_CSS.asp"-->
<%sql="select pic,showname from download where ID=" &request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql,conndown,1,1%>
<img src="<%=rs("pic")%>" border="0" alt=<%=rs("showname")%>Ïà¹ØÍ¼Æ¬>
<%rs.close
set rs=nothing
call footer()%>