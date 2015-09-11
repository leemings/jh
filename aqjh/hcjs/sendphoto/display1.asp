<!--#include file="connpic.asp"-->
<% 
id=request("id")
set rs=server.createobject("ADODB.recordset") 
If Session("aqjh_grade")=10 then
sql="select * from pic where id=" & id 
else
sql="select * from pic where 批准=1 and id=" & id 
end if
rs.open sql,conn,1,1
if rs.recordcount=0 then
response.write "<script language=javascript>alert('警告：找不到记录');window.close();</script>"
response.end
end if
Response.ContentType = "image/jpeg" 
Response.BinaryWrite rs("big")
rs.close 
set rs=nothing 
set conn=nothing 
%>