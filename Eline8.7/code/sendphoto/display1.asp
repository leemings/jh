<!--#include file="connpic.asp"-->
<% 
id=request("id")
set rs=server.createobject("ADODB.recordset") 

If Session("sjjh_grade")=10 then
sql="select * from pic where id=" & id 
else
sql="select * from pic where Åú×¼=1 and id=" & id 
end if
rs.open sql,conn,1,1 
Response.ContentType = "image/jpeg" 
Response.BinaryWrite rs("big")
rs.close 
set rs=nothing 
set conn=nothing 
%> 

