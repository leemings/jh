<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"


sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"


my=sjjh_name
money=abs(request.form("money"))
if money<>1000 and  money<>10000 and money<>100000 and money<>1000000 and money<>2000000 and money<>10000000 then 
	Response.Write "<script Language=Javascript>alert('������ʲô����');location.href = 'javascript:history.back()';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * from �û� where ����='"&sjjh_name&"'",conn
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
conn.execute "update �û� set �书=�书+"&int(money/1000)*15&",����=����-"& money &",����ʱ��=now() where ����='"&sjjh_name&"'"
rs.close
conn.close
set rs=nothing
set conn=nothing
Response.Write "<script Language=Javascript>alert('����ʦ���������п���ѧϰ������ʹ�Լ����书���,�书��+"& ((money/1000)*15) &"�㣬����������-"&money&"����');location.href = 'javascript:history.back()';;</script>"
%>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../bgcheetah.gif">