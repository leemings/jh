<%
if session("aqjh_name")<>Application("aqjh_user") then response.redirect "qqlist.asp"
if request("oicq")="" then response.redirect"send.asp?id=2"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="delete from QQ where oicq="&request("oicq")
rs.open sql,conn,3,2
conn.close
set conn=nothing
response.redirect "qqlist.asp?id=3"
%>
