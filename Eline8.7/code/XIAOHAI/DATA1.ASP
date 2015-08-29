<%
if Session("sjjh_name")="" then Response.Redirect "../errors.asp?id=016"

dbname="driver={Microsoft Access Driver (*.mdb)};DBQ="&Server.MapPath("Eline_Xiaohai.MDB")
set conn=server.createobject("adodb.connection")
conn.open dbname,"admin","chaopiaormb"
set rs=server.createobject("adodb.recordset")

%>