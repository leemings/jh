<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"


sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"


if sjjh_name="" then
	response.write("�Բ����㻹û��<a href=../../index.htm>��¼</a>")
else
	sex=request.form("sex")
if sex="" then
	response.write("�Բ����㻹û��<a href=../../index.htm>��¼</a>")
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM �û� where ����='"&sjjh_name&"'"&" and �Ա�= '" &sex&"'",conn
if rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=436"
else
if rs("����")<380 then
Response.Redirect "../../error.asp?id=437"
else
if day(rs("��Ǯ"))<day(now()) or month(rs("��Ǯ"))<month(now()) or year(rs("��Ǯ"))<year(now()) or isnull(rs("��Ǯ")) then
energy=rs("����")
tempdate=date
energytemp=energy+1000
sql="update �û� set ����=����-300,��Ǯ='"&tempdate&"',����='"&energytemp&"' where ����='"&sjjh_name&"'"
set rs=conn.execute(sql)
conn.close
set conn=nothing
set rs=nothing

Response.Redirect "../../ok.asp?id=601"
else
rs.close
Response.Redirect "../../error.asp?id=438"
end if
end if
end if
end if
end if
%>
