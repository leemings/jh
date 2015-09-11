<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
result=request("result")
id=request("id")
bg=request("bg")
my=request("my")
if result="罚500万" then
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rst=server.CreateObject ("adodb.recordset")
sql="update 用户 set 银两=银两-5000000 where 姓名='"&bg&"'"
conn.Execute(sql)
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rst=server.CreateObject ("adodb.recordset")
sql="update 用户 set 银两=银两+5000000 where 姓名='"&my&"'"
conn.Execute(sql)
end if
if result="罚1000万" then
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rst=server.CreateObject ("adodb.recordset")
sql="update 用户 set 银两=银两-10000000 where 姓名='"&bg&"'"
conn.Execute(sql)
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rst=server.CreateObject ("adodb.recordset")
sql="update 用户 set 银两=银两+10000000 where 姓名='"&my&"'"
conn.Execute(sql)
end if
if result="罚5000万" then
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rst=server.CreateObject ("adodb.recordset")
sql="update 用户 set 银两=银两-50000000 where 姓名='"&bg&"'"
conn.Execute(sql)
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rst=server.CreateObject ("adodb.recordset")
sql="update 用户 set 银两=银两+50000000 where 姓名='"&my&"'"
conn.Execute(sql)
end if
sql="update sy set 结果='1' where id=" & id & ""
conn.execute sql
Response.Redirect "manage.asp"
%>