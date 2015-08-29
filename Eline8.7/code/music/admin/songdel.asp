<!--#include file="function.asp"-->
<%CheckAdmin1%>
<!--#include file="conn.asp"-->
<%
   dim sql 
   dim rs
   set rs=server.createobject("adodb.recordset")
   sql="delete from MusicList where id="&request.QueryString("ID")
   rs.open sql,conn,1,1
   conn.close
   set conn=nothing
classid=request("classid")
SClassid=request("Sclassid")
Nclassid=request("Nclassid")
page=request("page")
response.redirect "songlist.asp?classid="+classid+"&SClassid="+SClassid+"&Nclassid="+Nclassid+"&page="& page
%>



