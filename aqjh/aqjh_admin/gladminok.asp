<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
Response.Write "<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>"
grade=clng(Request.Form("grade"))
name=trim(Request.Form("name"))
if name=Application("aqjh_user") then
	Response.Write "<script Language=Javascript>alert('��ʾ����վ������10�������Ը���!');location.href = 'gladmin.asp';</script>"
	response.end
end if
if aqjh_name<>Application("aqjh_user") and grade=10 then 
	Response.Write "<script Language=Javascript>alert('��ʾ���㲻����վ��������������10������Ա!');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("aqjh_usermdb")
select case Request.Form("submit")
case "�޸�"
	mp=trim(Request.Form("mp"))
	sf=trim(Request.Form("sf"))
	if sf="����" then
	conn.Close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�������������������!');location.href = 'gladmin.asp';</script>"
	response.end
	end if
	conn.execute("update �û� set grade="&grade&",����='�ٸ�',���='"&sf&"' where ����='"&name&"'")
case "ɾ��"
	conn.execute("update �û� set grade=1,����='����',���='����' where ����='"&name&"'")
case "���"
	conn.execute("update �û� set grade="&grade&",����='�ٸ�',���='����' where ����='"&name&"'")
end select
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','�����¼','["&Request.Form("submit")&"]ID="&name&"["&grade&"]�Ĺ���Ȩ��')"
conn.Close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��ʾ���������!');location.href = 'gladmin.asp';</script>"
response.end
%>