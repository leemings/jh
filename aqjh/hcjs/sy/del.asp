<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
id=request("id")
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rs=server.CreateObject ("adodb.recordset")
sql="delete from sy where id=" & id & ""
conn.execute sql
conn.close
set conn=nothing
%>
<body background="../../2di.jpg" bgcolor="#CCCCCC" topmargin="0">

<p><a href="manage.asp">·µ»ØÇëÇó½±Àø</a></p><LINK href="../../css.css" rel=stylesheet>