<!--#include file="function.asp"-->
<%CheckAdmin3%>
<!--#include file="conn.asp"-->
<%
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conn.execute("delete * from [user] where id="&request.QueryString("id")) 
set rs=nothing
conn.close
set conn=nothing
response.redirect "UserList.asp"
%>


