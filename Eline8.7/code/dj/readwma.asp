<!--#include file="conn.asp"-->
<%  
IF Session("auth") = True THEN 
Response.ContentType = "application/x-javascript"  
Response.Expires = 0  
Response.Expiresabsolute = Now() - 1  
Response.AddHeader "pragma","no-cache"  
Response.AddHeader "cache-control","private"  
Response.CacheControl = "no-cache"  
Session("auth") = False 
%>
<!--#include file="const.asp"--><ASX version = "3.0">
<%
if request("songid")<>"" then
    set rs=server.createobject("adodb.recordset")
    id=request("songid")
	sql="select * from MusicDJ where id in (" & id & ")"
    rs.open sql,conn,1,3
    while not rs.eof
	LF_Path=rs("LF_Path")
	rs("Hits")=rs("Hits")+1
	source=LF_Path
%><Entry>
<title><%=MusicName%></title>
<author><%=websitename%></author>
<Copyright><%=websiteurl%></Copyright>
<ref href = "<%=djserver%><%=source%>"/>
<param name="Artist" value="<%=DJUser%>"/>
<param name="Album" value="<%=WebSiteName%>"/>
<param name="Title" value="<%=rs("Musicname")%>"/>
</Entry><%
rs.movenext
wend
rs.Close
set rs=nothing
end if
conn.close
set conn=nothing
%></ASX>
<%ELSE%>这些代码受版权保护，所有权利保留！E线江湖总站→DJ舞吧<a href="HTTP://dj.51eline.com">dj.51eline.com</a><%END IF%>