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
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>
<title>���ɻ���</title>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../bg.gif">
<%if aqjh_name="" then Response.Redirect "../error.asp?id=210"
my=aqjh_name
money=abs(request.form("money"))
if money<>1000 and  money<>10000 and money<>100000 and money<>1000000 then
 	Response.Write "<script language=JavaScript>{alert('���ؾ��棬��Ҫ����');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=210"
	response.end
end if
if trim(rs("����"))="����" or trim(rs("����"))="��" or trim(rs("����"))="δ֪"  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('������������û���Լ������ɣ������ʲôѽ��');location.href = 'javascript:history.back()';</script>"
	response.end
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<5 then
	s=5-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�Բ���ϵͳ���ƣ����["&s&"����]�ٲ�����');}</script>"
	Response.End
end if	
if rs("����")<money then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���Ǯ���󲻹�ѽ��');location.href = 'javascript:history.back()';</script>"
	response.end
end if
mp=rs("����")
conn.execute "update �û� set ����=����-"& money &",����=����+"&int(money/500)&",���ɻ���=���ɻ���+"&money &",����ʱ��=now() where ����='"&aqjh_name&"'"
conn.execute "update p set h=h+"& money &" where a='" & mp &"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red>�����ɻ���["&aqjh_name&"]</font><font color=green>���Լ��������о�����"&money&"��,�������ǣ�"&int(money/500)&"��,���еĵ��Ӷ���м�����!</font>"
call showchat(says)
Response.Write "<script Language=Javascript>alert('������������о�����"& money &"�����������ǣ�"&int(money/500)&"����еĵ��Ӷ���м�������');location.href = 'javascript:history.back()';</script>"
%>