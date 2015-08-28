
<%
set connt=server.createobject("adodb.connection")
path="dbq="+server.mappath("jiudian.mdb")+";driver={microsoft access driver (*.mdb)};"
connt.open path
%>