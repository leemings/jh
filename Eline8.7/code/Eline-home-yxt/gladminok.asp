<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
Response.Write "<body text='#000000' background='../jhimg/bk_hc3w.gif' link='#0000FF' vlink='#0000FF' alink='#0000FF'>"
grade=clng(Request.Form("grade"))
name=trim(Request.Form("name"))
if name=Application("sjjh_user") then
	Response.Write "<script Language=Javascript>alert('��ʾ����վ������10�������Ը���!');location.href = 'gladmin.asp';</script>"
	response.end
end if
if sjjh_name<>Application("sjjh_user") and grade=10 then 
	Response.Write "<script Language=Javascript>alert('��ʾ���㲻����վ��������������10������Ա!');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("sjjh_usermdb")
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
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','�����¼','["&Request.Form("submit")&"]ID="&name&"["&grade&"]�Ĺ���Ȩ��')"
conn.Close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��ʾ���������!');location.href = 'gladmin.asp';</script>"
response.end
%>
