<%
   dim conn   
   dim connstr
   on error resume next
   connstr=Application("sjjh_usermdb")
   set conn=server.createobject("ADODB.CONNECTION")
   conn.open connstr 
%>

