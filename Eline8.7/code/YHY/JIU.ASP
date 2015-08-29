<%
set connt=server.createobject("adodb.connection")
path="dbq="+server.mappath("Eline_yhy.mdb")+";driver={microsoft access driver (*.mdb)};"
connt.open path
%>