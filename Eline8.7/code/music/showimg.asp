<%

Set conn=Server.CreateObject("ADODB.CONNECTION")

msgdata="DBQ="+server.mappath("photo/up/imagesa.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"

conn.open msgdata

set rec=server.createobject("ADODB.recordset")

strsql="select img from images where id=" & trim(request("id"))

rec.open strsql,msgdata,1,1

Response.ContentType = "image/*"

Response.BinaryWrite rec("img").getChunk(7500000)

rec.close

set rec=nothing

set conn=nothing

%>
