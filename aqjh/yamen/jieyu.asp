<%@ LANGUAGE=VBScript codepage ="936" %>
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
name=request("name")
my=aqjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="select * from �û� where ����='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
response.write "�㲻�ǽ������ˣ����ܾ���"
conn.close
response.end
else
if rs("�ȼ�")<30 then
response.write my & "����������Ҳ��������Ŷ�û�С�"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if rs("�Ṧ")<5000 then
response.write my & "���Ṧ��û����5000���ѧ��ȥ����������ѽ��"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if rs("�书")>2000 and (rs("״̬")<>"��" and rs("״̬")<>"��") then
randomize timer
r=int(rnd*3)
if r=1 then
sql="update �û� set ״̬='����' where ����='" & name & "'"
conn.execute sql
conn.close
Response.Redirect "laofan.asp"
else
sql="update �û� set ״̬='��' where ����='" & my & "'"
conn.execute sql
conn.close
response.write "�ٻ񲻳ɹ�����Ҳ��ץ��"
response.end
end if
else
response.write "�㲻�ܽ���"
conn.close
response.end
end if
end if
%><head></head>