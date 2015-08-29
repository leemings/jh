<%
   dim conn   
   dim connstr

   'on error resume next
   connstr="DBQ="+server.mappath("e_flash1.0.asp")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
   set conn=server.createobject("ADODB.CONNECTION")
   conn.open connstr 
%>

