<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
name=request("name")
my=aqjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	response.write "�㲻�ǽ������ˣ����ܾ���"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if rs("����")<50000000 then
response.write my & "�����������"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
rs.close
rs.open "select * from �û� where ����='" & name & "'",conn
if rs("״̬")="���" then
	response.write name & "�û��Ǳ�����Ĳ����ͷ�"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
conn.execute "update �û� set ����=����-1000000 where ����='"&aqjh_name&"'"
conn.execute "update �û� set ״̬='����',��¼=now() where ����='" & name & "'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "laofan.asp"
%>
