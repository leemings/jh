<!--#include file="conn.asp"-->
<%  IF Session("auth") = True THEN 
Response.ContentType = "application/x-javascript"  
Response.Expires = 0  
Response.Expiresabsolute = Now() - 1  
Response.AddHeader "pragma","no-cache"  
Response.AddHeader "cache-control","private"  
Response.CacheControl = "no-cache"  
Session("auth") = False 
%><!--#include file="const.asp"--><%
if request("songid")<>"" then
    set rs=server.createobject("adodb.recordset")
    id=request("songid")
	sql="select * from MusicDJ where id in (" & id & ")"
    rs.open sql,conn,1,3
    while not rs.eof
	LF_Path=rs("LF_Path")
	rs("Hits")=rs("Hits")+1
	source=LF_Path
%><%=djserver%><%=source%><%=chr(10)%><%
rs.movenext
wend
rs.Close
set rs=nothing
end if
conn.close
set conn=nothing
%>
<%ELSE%>��Щ�����ܰ�Ȩ����������Ȩ�����������ֽ�����վ��DJ���<a href="HTTP://dj.happyjh.com">dj.happyjh.com</a><p>��</p>
<%END IF%>