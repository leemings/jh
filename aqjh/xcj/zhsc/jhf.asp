<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../showchat.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');window.close();</script>"
	Response.End
end if
%>
<title>��ת������</title>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../../bg.gif">
<%my=aqjh_name
input=request.form("input")
if not isnumeric(input) then 
	Response.Write "<script language=JavaScript>{alert('��ʾ��["&input&"]���������ʹ�����֣�');location.href = 'flzh.asp';}</script>"
	Response.End 
end if
input=int(abs(input))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ��Ա��,��Ա�ȼ�,�ȼ�,����,������ from �û� where ����='"&aqjh_name&"'",conn,2,2
hy=rs("��Ա�ȼ�")
if rs("��Ա��")<input then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('��û����ô��Ľ��޷�ת����');location.href = 'flzh.asp';}</script>"
	response.end
end if
if rs("�ȼ�")<18 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('��ȼ�����18������ת����');location.href = 'flzh.asp';}</script>"
	response.end
end if
if input<1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('����������ûǮ���뻻��');location.href = 'flzh.asp';}</script>"
	response.end
end if
if (rs("����")+rs("������"))>20000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('���ܷ�������2000���Ѿ������ˣ�');location.href = 'flzh.asp';}</script>"
	response.end
end if
conn.execute "update �û� set ��Ա��=��Ա��-" & input & ",����=����+"& int(input*5000)&" where ����='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red>��ת���г���["&aqjh_name&"]</font><font color=blue>��"&input&"Ԫ��ת������"&int(input*5000)&"�㷨��,���Ƿ�������!</font>"
call showchat(says)
Response.Write "<script Language=Javascript>alert('"&aqjh_name&"��ת������:"&input*5000&"�㣡');location.href = 'flzh.asp';</script>"
%>