<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../showchat.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<title>ְҵת��</title>
<body bgcolor="#CCCCCC">
<%ziye=trim(request.form("jiu"))
if len(ziye)>3 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ҫ��ʲô����');window.close();}</script>"
	Response.End 
end if
if InStr(ziye,"or")<>0 or InStr(ziye,"'")<>0 or InStr(ziye,"`")<>0 or InStr(ziye,"=")<>0 or InStr(ziye,"-")<>0 or InStr(ziye,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');window.close();}</script>"
	Response.End 
end if
if ziye<>"�ɱ�" and ziye<>"�ɿ�" and ziye<>"��ľ" and ziye<>"ð�ռ�" and ziye<>"����ʦ" and ziye<>"�з�" and ziye<>"Ů��" and ziye<>"��ʦ" and ziye<>"���" and ziye<>"С͵" and ziye<>"����ʦ" and ziye<>"����ʦ" and ziye<>"ħ��ʦ" and ziye<>"��ҹ��" and ziye<>"���" then
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');window.close();}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='"&aqjh_name&"'",conn,2,2
if rs("ְҵ")=ziye then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ԭ���������ְҵ����תʲôѽ��');window.close();}</script>"
	response.end
end if

if ziye="����ʦ" and (aqjh_jhdj<20 or aqjh_grade>=11 or pi="����") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ����תְΪ����ʦ������ս���ȼ����Ϊ20���������Ƿǳ����ˣ��ǹٸ���Ա��');window.close();}</script>"
		response.end
end if
if ziye="��ҹ��" and (aqjh_jhdj<15 or aqjh_grade>=11 or pi="����") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ����תְΪ��ҹ�㣬������ս���ȼ����Ϊ15���������Ƿǳ����ˣ��ǹٸ���Ա��');window.close();}</script>"
		response.end
end if
if ziye="С͵" and (aqjh_jhdj<20 or aqjh_grade>=6 or pi="����") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ����תְΪС͵��������ս���ȼ����Ϊ20���������Ƿǳ����ˣ��ǹٸ���Ա��');window.close();}</script>"
		response.end
end if
if ziye="ð�ռ�" and (aqjh_jhdj<50 or pi="����") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ����תְΪð�ռң�������ս���ȼ����Ϊ50���������Ƿǳ����ˣ�');window.close();}</script>"
		response.end
end if
if ziye="����ʦ" and (aqjh_jhdj<60 or pi="����" or pi="����") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ����תְΪ����ʦ��������ս���ȼ����Ϊ60���������Ƿǳ����ˣ�');window.close();}</script>"
		response.end
end if
if ziye="ħ��ʦ" and hy=rs("��Ա�ȼ�") and (aqjh_jhdj<80 or pi="����" or hy<1) then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ����תְΪħ��ʦ�������ǻ�Ա��ս���ȼ����Ϊ80���������Ƿǳ����ˣ�');window.close();}</script>"
		response.end
end if
pi=rs("����")
sex=rs("�Ա�")
if ziye="�з�" and (aqjh_jhdj<20 or aqjh_grade>=6 or pi="����" or sex<>"��") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����תְΪ�з����������еģ�ս���ȼ����Ϊ20���������Ƿǳ����ˣ��ǹٸ���Ա��');window.close();}</script>"
		response.end
		end if
		if ziye="Ů��" and (aqjh_jhdj<20 or aqjh_grade>=6 or pi="����" or sex<>"Ů") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ����תְΪŮ����������Ů�ģ�ս���ȼ����Ϊ20���������Ƿǳ����ˣ��ǹٸ���Ա��');window.close();}</script>"
		response.end
		end if
if rs("����")<500000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����Ǯ����50��');window.close();}</script>"
	response.end
end if
if rs("���")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������10����Ҷ�û�У�����Ҫְҵ��');window.close();}</script>"
	response.end
end if
if rs("����")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��ת��ְҵ��Ҫ1000�����ĸ�����ȥ��߰ɣ�');window.close();}</script>"
	response.end
end if
pi=rs("����")
sex=rs("�Ա�")
if ziye="�ɱ�" or ziye="�ɿ�" or ziye="��ľ" or ziye="�з�" or ziye="Ů��" or ziye="���" or ziye="��ʦ" or ziye="����ʦ" or ziye="����ʦ" or ziye="��ҹ��" or ziye="���" or ziye="С͵"  then
	conn.execute "update �û� set ����=����-100,���=���-5,����=����-500000,ְҵ='"& ziye &"' where ����='"&aqjh_name&"'",conn
	else
end if
if ziye="ħ��ʦ" or ziye="ð�ռ�" or ziye="����ʦ"  then
	conn.execute "update �û� set ����=����-500,���=���-10,����=����-500000,ְҵ='"& ziye &"' where ����='"&aqjh_name&"'",conn
	else
end if
rs.close
set rs=nothing
'conn.close
set conn=nothing
says="<font color=red>��ְҵת����["&aqjh_name&"]</font><font color=blue>ת����ְҵ<font color=red>["&ziye&"]</font>,ϣ�����ְҵ���������ģ�</font>"
call showchat(says)
Response.Write "<script language=JavaScript>{alert('��ʾ����ְҵת�����ˣ�"& ziye &"���������ȷ���رյ�ǰ���ڣ�');window.close();}</script>"
response.end
%>