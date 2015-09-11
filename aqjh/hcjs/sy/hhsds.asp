<%
number=request("number")+0
if number=0 then number=14
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rs=server.CreateObject ("adodb.recordset")
sql="select * from sy order by id DESC"
set rs=conn.execute(sql)
i=0
do while not rs.eof and i<number
%>
document.write("<li><a href=javaScript:MM_openBrWindow('sy/rootdisp.asp?id=<%=rs("id")%>')><%=rs("标题")%></a><%if rs("结果")="N" then%>未审<% end if %><%if  rs("结果")<>"N" then %>已审<% end if %><br>");
<%'	response.write rs("title") & "<br>"
	i=i+1
	rs.movenext
loop
set rs=nothing
conn.close
set conn=nothing
%>