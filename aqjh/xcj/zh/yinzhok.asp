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
<title>�ֽ�ת�����</title>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../../bg.gif">
<%my=aqjh_name
input=request.form("input1")
if not isnumeric(input) then 
	Response.Write "<script language=JavaScript>{alert('��ʾ��["&input&"]���������ʹ�����֣�');location.href = 'yinzh.ASP';}</script>"
	Response.End 
end if
input=int(abs(input))
if input/1<>int(input/1) then
 	Response.Write "<script language=JavaScript>{alert('������1������������ֵ��');location.href = 'yule2.asp';}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ���,����,���,�ȼ� from �û� where ����='" & aqjh_name & "'",conn,2,2
if rs("���")<input then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('��ʾ����û����ô��Ľ���޷�ת����');location.href = 'yule2.asp';}</script>"
	response.end
end if
if rs("�ȼ�")<25 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('��ȼ�����25������ת����');location.href = 'yinzh.asp';}</script>"
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
if (rs("����")+rs("���"))>2000000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('��Ǯ�Ѿ�����20���ˣ��Ѿ������ˣ�');location.href = 'flzh.asp';}</script>"
	response.end
end if
conn.execute "update �û� set ���=���-"&input&",����=����+" & int(input*1000000) & " where ����='"&aqjh_name&"'"
conn.close
set conn=nothing
says="<font color=red>��ת���г���["&aqjh_name&"]</font><font color=blue>��"&input&"����Ҷ��ֳ���"&int(input*1000000)&"�������ɴ�����!</font>"
call showchat(says)
Response.Write "<script Language=Javascript>alert('"&aqjh_name&"��"&input&"�����ת����������:"& int(input*1000000)&"����');location.href = 'yinzh.asp';</script>"
%>
