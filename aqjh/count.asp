<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM sm WHERE a='¼ÇÊýÆ÷'",conn
Application("count")=rs("b")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
document.write('<%=Application("count")%>');