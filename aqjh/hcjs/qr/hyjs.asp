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
rs.open "select ����,�Ա�,����,hw from �û� where ����='"&aqjh_name&"'",conn
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
rs.close
if hai="����" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�㻹û�������أ�����Ҫ��飡');window.close();}</script>"
	Response.End
end if
if yin<1000000 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���һ����ȡ����100�����Ǯ������');window.close();}</script>"
	Response.End
end if
hdata=split(hai,"|")
baby=hdata(0)
babysj=hdata(1)
babytl=clng(hdata(2))
babyzt=hdata(3)
erase hdata
sysj=DateDiff("d",babysj,now())
top="�������㻳�е�"&sysj&"�죬���̥��һ��������һ��Ҫע�Ᵽ������Ӵ��"
if sysj>=3 and sysj<10 and babyzt<>"ҩ1" then
	babyzt="ҩ"
	top="������ǻ��еĵ�"&sysj&"�죬��������Ҫȥ���ñ�̥ҩ�ˡ�һ���շ�500��������15��ҡ�"
end if
if sysj>=10 and sysj<20 and baby<>"��1" then
	babyzt="��"
	top="������ǻ��еĵ�"&sysj&"�죬��Ҫȥ��̥���ˡ�һ���շ�1500��������20��ҡ�"
end if
if sysj=20 then
	babyzt="��"
	top="��ѽ�������Ͼ�Ҫ���ˣ���ȥ�����ɡ�����ҽԺ�����˵����壬����С�������κη��á�"
end if
if sysj=21 then
	babyzt="��"
	top="ι���㶼��21���ˣ���ʱ���ˣ�����������������������Ӥ����̥�������ˡ����ȥҽԺ�ɡ�"
end if
newbl=baby&"|"&babysj&"|"&babytl&"|"&babyzt
if sysj>21 then
	newbl="����"
	top="�������̥���ڸ��г���21��δ�������Ѿ�̥�������ˡ�������Ϊ����δ������С��������ʧ��Ĭ���ɡ���"
end if
conn.execute "update �û� set ����='"&newbl&"',����=����-1000000 where ����='"&aqjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>���м��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<div align="center">
  <p>&nbsp;</p>
  <p><font size="7" color="#FF0000">������������-���м��</font></p>
  <p>&nbsp;</p>
  <p><img src="images/zd.gif" width="300" height="144"> <br>
    <font size="5"><%=top%></font></p>
  <p><font size="5"><a href="fkyy.asp">����</a></font></p>
</div>
</body></html>