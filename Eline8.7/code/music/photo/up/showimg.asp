<!--#include file="conn.asp"-->

<%



set rec=server.createobject("ADODB.recordset")

strsql="select img from images where id=" & trim(request("id"))

rec.open strsql,msgdata,1,1

Response.ContentType = "image/*"

Response.BinaryWrite rec("img").getChunk(7500000)

rec.close

set rec=nothing

set conn=nothing

%>
