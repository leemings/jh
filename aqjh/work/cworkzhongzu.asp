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
%>
<title>����ת��</title>
<body bgcolor="#CCCCCC">
<%zhongzu=trim(request.form("jiu"))
if len(zhongzu)>3 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ҫ��ʲô����');window.close();}</script>"
	Response.End 
end if
if InStr(zhongzu,"or")<>0 or InStr(zhongzu,"'")<>0 or InStr(zhongzu,"`")<>0 or InStr(zhongzu,"=")<>0 or InStr(zhongzu,"-")<>0 or InStr(zhongzu,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');window.close();}</script>"
	Response.End 
end if
if zhongzu<>"����" and zhongzu<>"����" and zhongzu<>"ħ��"  then
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');window.close();}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='"&aqjh_name&"'",conn,2,2
if rs("����")=zhongzu then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ԭ������������壬��תʲôѽ��');window.close();}</script>"
	response.end
end if

if zhongzu="����" and (aqjh_jhdj<150 or aqjh_grade>=11 or pi="����") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ����תְΪ���壬����ս���ȼ����Ϊ150���������Ƿǳ�����');window.close();}</script>"
		response.end
end if
if zhongzu="����" and (aqjh_jhdj<90 or aqjh_grade>=11 or pi="����" ) then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ����תְΪ���壬������ս���ȼ����Ϊ90���������Ƿǳ�����');window.close();}</script>"
		response.end
end if
if zhongzu="ħ��" and (aqjh_jhdj<100 or aqjh_grade>=11 or pi="����" ) then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ����תְΪħ�壬������ս���ȼ����Ϊ100���������Ƿǳ�����');window.close();}</script>"
		response.end
		end if
if rs("����")<5000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����Ǯ����500��');window.close();}</script>"
	response.end
end if
if rs("����")<2 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������ת����Ҫ����2�㣡');window.close();}</script>"
	response.end
end if
if rs("ת��")<3 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������ת������ת��3�����ϣ�');window.close();}</script>"
	response.end
end if
pi=rs("����")
sex=rs("�Ա�")
if zhongzu="����" or zhongzu="����" or zhongzu="ħ��"  then
	conn.execute "update �û� set ����=����-50000000,����=����-1,����='"& zhongzu &"' where ����='"&aqjh_name&"'",conn
	else
end if
rs.close
set rs=nothing
'conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('��ʾ����ת�����ˣ�"& zhongzu &"���壬�����۳�һ�㣬���ȷ���رյ�ǰ���ڣ�');window.close();}</script>"
response.end
%>