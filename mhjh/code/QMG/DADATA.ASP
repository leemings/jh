<%
set connt=server.createobject("adodb.connection")
path="dbq="+server.mappath("mian.mdb")+";driver={microsoft access driver (*.mdb)};"
connt.open path
%>
