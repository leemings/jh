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
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');window.close();</script>"
	Response.End
end if
my=aqjh_name
money=abs(request.form("money"))
if money<>1000 and  money<>10000 and money<>100000 and money<>1000000 and money<>2000000 and money<>10000000 then 
	Response.Write "<script Language=Javascript>alert('������ʲô����');location.href = 'javascript:history.back()';</script>"
	response.end
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
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<15 then
	s=15-sj
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
	Response.Write "<script Language=Javascript>alert('������������ѽ����Ȼ��ʦͽ������û��Ǯ�������еģ�');location.href = 'javascript:history.back()';;</script>"
	response.end
end if
conn.execute "update �û� set �书=�书+"&int(money/1000)*15&",����=����-"& money &",����ʱ��=now() where ����='"&aqjh_name&"'"
rs.close
conn.close
set rs=nothing
set conn=nothing
Response.Write "<script Language=Javascript>alert('����ʦ���������п���ѧϰ������ʹ�Լ����书���,�书��+"& ((money/1000)*15) &"�㣬����������-"&money&"����');location.href = 'javascript:history.back()';;</script>"
%>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../bg.gif">