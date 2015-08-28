<%
set conntt=server.createobject("adodb.connection")
path="dbq="+server.mappath("../21ex/fangzi.asp")+";driver={microsoft access driver (*.mdb)};"
conntt.open path
%>
