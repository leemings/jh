<!--#include file=conn.asp-->
<%
response.buffer=true
Response.Cookies("skin")("skinid_0")=""
dim rs
set rs=conn.execute("select boardid from board order by rootid ,orders")
do while not rs.eof
Response.Cookies("skin")("skinid_"&rs(0))=""
rs.movenext
loop
set rs=nothing
conn.close
set conn=nothing
response.redirect "index.asp"
%>
