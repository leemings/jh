<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"

lb=lcase(trim(request("wqlb")))
if lb="" or (lb<>"z1" and lb<>"z2" and lb<>"z3" and lb<>"z4" and lb<>"z5" and lb<>"z6") then
	Response.Write "<script Language=Javascript>alert('��Ʒ��������');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if InStr(lb,"or")<>0 or InStr(lb,"=")<>0 or InStr(lb,"`")<>0 or InStr(lb,"'")<>0 or InStr(lb," ")<>0 or InStr(lb,"��")<>0 or InStr(lb,"'")<>0 or InStr(lb,chr(34))<>0 or InStr(lb,"\")<>0 or InStr(lb,",")<>0 or InStr(lb,"<")<>0 or InStr(lb,">")<>0 then Response.Redirect "../../error.asp?id=54"
select case lb
	case z1
		lbmc="ͷ��"
	case z2
		lbmc="װ��"
	case z3
		lbmc="����"
	case z4
		lbmc="����"
	case z5
		lbmc="����"
	case z6
		lbmc="˫��"
end select
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT "&lb&" FROM �û� WHERE ����='"&aqjh_name&"'",conn
nnn=trim(rs(lb))
if trim(nnn)="" or trim(nnn)=null then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����["&lbmc&"]��û���κ�װ��������ʲô��������');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
data=split(nnn,"|")
ygj=data(2)
yfy=data(3)
conn.execute "update �û� set ������=������-"&ygj&",������=������-"&yfy&" where ����='"&aqjh_name&"'"
conn.execute "update �û� set "&lb&"='' where ����='"&aqjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>����װ��</title>
</head>
<body background="back.gif">
<p align="center"><font size="6" color="#FF0000"><b>����װ��</b></font></p>
<p align="center">�㶪����<font color=green><%=lbmc%></font>��װ��<font color=red>[<%=data(0)%>]</font>����Ҫ����װ�����빺���µ�װ����<br>
  <br>
  <a href="index.asp">����</a> <br>
  <br>
  <font size="2">��Ȩ���С����齭����վ��</font> </p>
</body></html>