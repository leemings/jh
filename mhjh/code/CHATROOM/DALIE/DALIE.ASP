<%
set connt=server.createobject("adodb.connection")
path="dbq="+server.mappath("dalie.mdb")+";driver={microsoft access driver (*.mdb)};"
connt.open path
%>