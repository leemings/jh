
<!--#include file="conn.asp"-->
<%
response.buffer=true
formsize=request.totalbytes
formdata=request.binaryread(formsize)
bncrlf=chrB(13) & chrB(10)
divider=leftB(formdata,clng(instrb(formdata,bncrlf))-1)
datastart=instrb(formdata,bncrlf & bncrlf)+4
dataend=instrb(datastart+1,formdata,divider)-datastart
mydata=midb(formdata,datastart,dataend)

set rec=server.createobject("ADODB.recordset")
sql= "SELECT * FROM images where id is null"
rec.open sql,conn,1,3

rec.addnew
rec("img").appendchunk mydata
rec.update
rec.close
set rec=nothing
set conn=nothing
%>
<meta http-equiv="refresh" content="0;URL=Yxpic.asp">