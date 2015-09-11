<%
set connt=server.createobject("adodb.connection")
path="dbq="+server.mappath("../../aqjh_data/pig.asp")+";driver={microsoft access driver (*.mdb)};"
connt.open path
%>