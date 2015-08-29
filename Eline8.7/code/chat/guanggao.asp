<%@ LANGUAGE=VBScript codepage ="936" %>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM sm where a='¹ã¸æ'",conn,2,2
%>
<html>
<%=rs("c")%>
<%rs.close 
set rs=nothing
conn.close
set rs=nothing
%>
</html>