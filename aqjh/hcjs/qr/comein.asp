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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,���,�Ա�,����,hw from �û� where ����='"&aqjh_name&"'",conn
if rs("�Ա�")="��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('����ҽԺ������ֹ����');window.close();}</script>"
	Response.End
end if
hai=rs("����")
yin=rs("����")
jb=rs("���")
rs.close
if hai="����" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�㻹û�������أ�����ʲô̥ѽ��');window.close();}</script>"
	Response.End
end if
hdata=split(hai,"|")
baby=hdata(0)
babysj=hdata(1)
babytl=clng(hdata(2))
babyzt=hdata(3)
erase hdata
if babyzt="����" or babyzt="ҩ1" or babyzt="��1" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���̥��һ����������ʱ����Ҫ�����κα�����ʩ��');window.close();}</script>"
	Response.End
end if
if babyzt="��" or babyzt="��" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ѽ�����Ҫ���ˣ������ڸ��ڣ��쵽����ȥѽ��');window.close();}</script>"
	Response.End
end if
if babyzt="ҩ" then
	yin1=5000000
	jb1=15
	if yin<yin1 or jb<jb1 then
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('����Ҫʹ�ñ�̥ҩ�����ã�����500�򣬽��15�������Ǯ������');window.close();}</script>"
		Response.End
	end if
	babyzt="ҩ1"
	babytl=babytl+300
	h="<img src=images/yao.gif width=100 height=165><br><br>ҽ����������˱�̥ҩ��̥��������������300�㡣�շ�500��15��ҡ�"
end if
if babyzt="��" then
	yin1=15000000
	jb1=20
	if yin<yin1 or jb<jb1 then
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('����Ҫ��̥�룬���ã�����1500�򣬽��20�������Ǯ������');window.close();}</script>"
		Response.End
	end if
	babyzt="��1"
	babytl=babytl+900
	h="<img src=images/z.gif width=100 height=165><br><br>ҽ��Ϊ�����һ֧��̥�룬̥��������������900�㣬�շ�����1500��20��ҡ�"
end if
newbl=baby&"|"&babysj&"|"&babytl&"|"&babyzt
conn.execute "update �û� set ����=����-"&yin1&",���=���-"&jb1&",����='"&newbl&"' where ����='"&aqjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000">
 
<div align="center"><br>
  <br>
  <font size="6" color="#FF0000">��������_������</font><br>
  <br>
  <br>
  <%=h%></div>
</body>
</html>