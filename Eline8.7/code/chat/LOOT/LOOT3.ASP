<%
set connt=server.createobject("adodb.connection")
path="dbq="+server.mappath("Eline_loot.mdb")+";driver={microsoft access driver (*.mdb)};"
connt.open path
%>