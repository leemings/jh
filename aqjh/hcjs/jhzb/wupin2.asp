<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
name=lcase(trim(request("name")))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT ����ʱ��,w3,z1,z2,z3,z4,z5,z6 FROM �û� WHERE ����='"&aqjh_name&"'",conn,3,3
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�Բ���ϵͳ���ƣ����["&s&"����]�ٲ�����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
duyao=abate(rs("w3"),name,1)
z11=rs("z1")
z21=rs("z2")
z31=rs("z3")
z41=rs("z4")
z51=rs("z5")
z61=rs("z6")
rs.close
rs.open "select b,f,g,j,k FROM b WHERE a='" & name &"'",conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������Ʒ�����ݿ��в����ڣ�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
nai=rs("j")
lb1=rs("b")
wpfy=rs("g")
wpgj=rs("f")
wptx=rs("k")
select case lb1
	case "ͷ��"
		lb="z1"
		zb=z11
	case "װ��"
		lb="z2"
		zb=z21
	case "����"
		lb="z3"
		zb=z31
	case "����"
		lb="z4"
		zb=z41
	case "����","����"
		lb="z5"
		zb=z51
	case "˫��"
		lb="z6"
		zb=z61
	case else
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ��������ݳ�������򿪷�����ϵ��');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End
end select
if trim(zb)<>"" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���["&lb1&"]װ������������Ҫ�������������ɵ�������װ���µģ�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
wh=0
conn.execute "update �û� set ������=������+"& wpgj &",������=������+"&wpfy&" where ����='"&aqjh_name&"'"
conn.execute "update �û� set w3='"&duyao&"',"&lb&"='"&name&"|"&nai&"|"&wpgj&"|"&wpfy&"|"&wptx&"|"&wh&"' where  ����='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��ʾ��װ����Ʒ["&name&"]���\n��ȷ������');location.href = 'javascript:history.go(-1)';</script>"
response.end
%>