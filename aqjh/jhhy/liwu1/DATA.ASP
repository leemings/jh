<%
set connt=server.createobject("adodb.connection")
path="dbq="+server.mappath("sdlw.asp")+";driver={microsoft access driver (*.mdb)};"
connt.open path
%>