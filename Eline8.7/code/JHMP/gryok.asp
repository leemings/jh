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


%>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../bgcheetah.gif">
<%my=sjjh_name
money=abs(request.form("money"))
if money<>1000 and  money<>10000 and money<>100000 and money<>1000000 then 
	Response.Write "<script Language=Javascript>alert('������ʲô����');window.close();</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * from �û� where ����='"&sjjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	coon.close
	set conn=nothing
	Response.Redirect "../error.asp?id=210"
	response.end
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<10 then
	s=10-sj
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
	Response.Write "<script Language=Javascript>alert('��û��Ǯѽ��������ʲô����');window.close();</script>"
	response.end
end if
conn.execute "update �û� set ����=����+"&int(money/500)&",����=����-"& money &",����ʱ��=now() where ����='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('������Ĺ¶�Ժ������"& money &"������������"&int(money/500)&"�㣡');window.close();</script>"
rs.close
conn.close
set rs=nothing
set conn=nothing
%>
