<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
name=request("name")
my=sjjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * from �û� where ����='"&sjjh_name&"'",conn
if rs.eof or rs.bof then
	response.write "�㲻�ǽ������ˣ����ܾ���"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if rs("����")<10000000 then
response.write my & "�������ҲѧȥȥҽԺ������"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
rs.close
rs.open "select * from �û� where ����='" & name & "'",conn
if rs("״̬")="���" then
	response.write name & "�����ط�,�����ܣ�"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if rs("״̬")="��" then
	response.write name & "����ҽԺ��������ط��˰ɣ�"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if rs("״̬")="��" then
	response.write name & "����ҽԺ��������ط��˰ɣ�"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if rs("״̬")="��" then
	response.write name & "����ҽԺ��������ط��˰ɣ�"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
conn.execute "update �û� set ����=����-10000000 where ����='"&sjjh_name&"'"
conn.execute "update �û� set ״̬='����',��¼=now() where ����='" & name & "'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "yiyuan.asp"
%>
