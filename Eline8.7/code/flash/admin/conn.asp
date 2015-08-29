<%
   Server.ScriptTimeOut=150
   dim conn   
   dim connstr

   'on error resume next
   connstr="DBQ="+server.mappath("eline_yxt.asa")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
   set conn=server.createobject("ADODB.CONNECTION")
   conn.open connstr 
%>

