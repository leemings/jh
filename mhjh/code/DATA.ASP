<%
set connt=server.createobject("adodb.connection")
path="dbq="+server.mappath("21ex/jihe.asp")+";driver={microsoft access driver (*.mdb)};"
connt.open path
%>