<%
   dim conn   
   dim connstr
   on error resume next
   connstr="DBQ="+server.mappath("../aqjh_data/aqjh.asp")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
   set conn=server.createobject("ADODB.CONNECTION")
   conn.open connstr 
%>