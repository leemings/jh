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
if result="��500��" then
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rst=server.CreateObject ("adodb.recordset")
sql="update �û� set ����=����-5000000 where ����='"&bg&"'"
conn.Execute(sql)
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rst=server.CreateObject ("adodb.recordset")
sql="update �û� set ����=����+5000000 where ����='"&my&"'"
conn.Execute(sql)
end if
if result="��1000��" then
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rst=server.CreateObject ("adodb.recordset")
sql="update �û� set ����=����-10000000 where ����='"&bg&"'"
conn.Execute(sql)
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rst=server.CreateObject ("adodb.recordset")
sql="update �û� set ����=����+10000000 where ����='"&my&"'"
conn.Execute(sql)
end if
if result="��5000��" then
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rst=server.CreateObject ("adodb.recordset")
sql="update �û� set ����=����-50000000 where ����='"&bg&"'"
conn.Execute(sql)
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rst=server.CreateObject ("adodb.recordset")
sql="update �û� set ����=����+50000000 where ����='"&my&"'"
conn.Execute(sql)
end if
sql="update sy set ���='1' where id=" & id & ""
conn.execute sql
Response.Redirect "manage.asp"
%>