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
<title>��ת���书</title>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../../bg.gif">
<%my=aqjh_name
input=request.form("input")
if not isnumeric(input) then 
	Response.Write "<script language=JavaScript>{alert('��ʾ��["&input&"]���������ʹ�����֣�');location.href = 'wgzh.asp';}</script>"
	Response.End 
end if
input=int(abs(input))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if Weekday(date())=7 and Hour(time())>=18 then
Response.Write "<script Language=Javascript>alert('��ʾ�����俿ƽʱ�����ᱦ֮�ս��ã�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
rs.open "select ���,��Ա�ȼ�,�书,�ȼ� from �û� where ����='"&aqjh_name&"'",conn,2,2
hy=rs("��Ա�ȼ�")
if rs("���")<input then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('��ʾ����û����ô��Ľ���޷�ת����');location.href = 'wgzh.asp';}</script>"
	response.end
end if
if rs("�ȼ�")<30 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('��ȼ�����30������ת����');location.href = 'wgzh.asp';}</script>"
	response.end
end if
if input<=1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('����������ûǮ���뻻��');location.href = 'wgzh.asp';}</script>"
	response.end
end if
if rs("�书")>6000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('���书�Ѿ�����600�򣬹��������ˣ�');location.href = 'wgzh.asp';}</script>"
	response.end
end if
conn.execute "update �û� set ���=���-" & input & ",�书=�书+"& int(input*1000)&" where ����='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red>��ת���г���["&aqjh_name&"]</font><font color=blue>��"&input&"Ԫ���ת������"&int(input*1000)&"���书,�����书����!</font>"
call showchat(says)
Response.Write "<script Language=Javascript>alert('"&aqjh_name&"��ת���书:"&input*1000&"�㣡');location.href = 'wgzh.asp';</script>"
%>