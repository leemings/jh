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
rs.open "select �Ա�,����,hw from �û� where ����='"&aqjh_name&"'",conn
if rs("�Ա�")="��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('����ҽԺ������ֹ����');window.close();}</script>"
	Response.End
end if
%>
<html>
<head>
<title>����ҽԺ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<div align="center"> 
  <p>&nbsp;</p>
  <p><font size="7" color="#800000">������������</font></p>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p><a href="comein.asp"><img src="images/sjyy.gif" width="161" height="145" border="0"></a></p>
  <p>&nbsp;</p>
  <p><font size="6"><a href="hyjs.asp">���м��</a> <a href="comein.asp">������</a> <a href="baby.asp">�� ��</a></font></p> 
</div>
</body>
</html>
