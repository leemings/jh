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
<title>��ȡ����</title>
<body bgcolor="#CCCCCC">
<%renwu=trim(request.form("jiu"))
if InStr(renwu,"or")<>0 or InStr(renwu,"'")<>0 or InStr(renwu,"`")<>0 or InStr(renwu,"=")<>0 or InStr(renwu,"-")<>0 or InStr(renwu,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');window.close();}</script>"
	Response.End 
end if
if renwu<>"��������" and renwu<>"��������" and renwu<>"ħ������"  then
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');window.close();}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='"&aqjh_name&"'",conn,2,2
if rs("����")=renwu then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ԭ������������񣬻���ʲôѽ��');window.close();}</script>"
	response.end
end if

if renwu="��������" and (aqjh_jhdj<120 or aqjh_grade>=11 or pi="����") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ���������񣬱���ս���ȼ����Ϊ120���������Ƿǳ�����');window.close();}</script>"
		response.end
end if
if renwu="��������" and (aqjh_jhdj<60 or aqjh_grade>=11 or pi="����") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ���������񣬱�����ս���ȼ����Ϊ60���������Ƿǳ�����');window.close();}</script>"
		response.end
end if
if renwu="ħ������" and (aqjh_jhdj<90 or aqjh_grade>=11 or pi="����") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ��ħ�����񣬱�����ս���ȼ����Ϊ90���������Ƿǳ�����');window.close();}</script>"
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
pi=rs("����")
sex=rs("�Ա�")
if renwu="��������" or renwu="��������" or renwu="ħ������"  then
	conn.execute "update �û� set ����=����-50000000,����='"& renwu &"',����ʱ��='"& time &"' where ����='"&aqjh_name&"'",conn
	else
end if
rs.close
set rs=nothing
'conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('��ʾ���������ˣ�"& renwu &"�����ȷ���رյ�ǰ���ڣ�');window.close();}</script>"
response.end
%>