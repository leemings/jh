<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0

myid= Replace(Session("aqjh_name"), "'", "''")
if myid="" then
	Response.Write("�������ʺš�")
	Response.End
End If

Set Conn = Server.CreateObject("ADODB.Connection") 
Conn.Open Application("aqjh_usermdb")
Set Rs = Server.CreateObject("ADODB.Recordset")

	rs.open "select * from �û� where ����='"& myid &"'",conn
	If rs.Bof and rs.Eof then
		info="�㲻�ǽ�������"
		err_num=err_num+1
	else
		jh_yl=rs("����")
	end if
	rs.close

	Sql="SELECT b.*, wpname.* FROM wpname LEFT OUTER JOIN b ON wpname.wp_name = b.a "
Sql=Sql&" WHERE wpname.wp_sl>0 and (wpname.wp_user = '"&myid&"') AND (b.b = '���' OR b.b = '���')"
	'Response.Write Sql
	'Response.End

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>����</title>

<link href="base.css" rel="stylesheet" type="text/css" />
</head>
<body> 
<div align="center">���е���Ʒ������<%= jh_yl %> <a href="index.asp" target="_parent">����ȥ</a> </div>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="3"> 
  <tr align="center" bgcolor="#FF9900"> 
    <td>��Ʒ��</td> 
    <td>��Ʒ����</td> 
    <td>˵��</td> 
    <td>ͼƬ</td> 
  </tr> 
  <%
  Rs.Open Sql,Conn,0,3
  Rs_i=0
	While (NOT Rs.EOF)
	wp_image=(Rs.Fields.Item("i").Value)
	if isnull(wp_image) or wp_image="" or wp_image="��" then  wp_image="0.gif"
	wp_name=Rs.Fields.Item("wp_name").Value
    %> 
  <tr align="center" bgcolor="#FFCC99"> 
    <td><%=(Rs.Fields.Item("wp_name").Value)%></td> 
    <td><%=(Rs.Fields.Item("wp_sl").Value)%></td> 
    <td><%=(Rs.Fields.Item("c").Value)%></td> 
    <td><img style = "FILTER: chroma(color:#FFFFFF)" src="images/<%= wp_image %>" width="30" height="30" /></td> 
  </tr> 
  <% 
  Rs_i=Rs_i+1
  Rs.MoveNext()
Wend
Rs.Close()
Conn.close
set Conn=nothing
%> 
</table> 
</body>
</html>

</body>
</html>



