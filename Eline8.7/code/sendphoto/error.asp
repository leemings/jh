<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
id=Trim(Request.QueryString("id"))
chatroombgimage=Application("hxf_c_chatroombgimage")
chatroombgcolor=Application("hxf_c_chatroombgcolor")%><html>
<head>
<title>������ʾ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="chat/readonly/style.css">
</head>
<body bgcolor="<%if chatroombgcolor="" then
Response.Write "FFFFFF"
else
Response.Write chatroombgcolor
end if%>"<%if chatroombgimage<>"" then%> background=<%=chatroombgimage%><%end if%> leftmargin="0" topmargin="0">
<table width="100%" border="0" height="100%">
<tr align="center"> 
<td>
<form method="post" action="">
<table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" bgcolor="E0E0E0">
<tr>
<td>
              <table border="0" bgcolor="#0099FF" cellspacing="0" cellpadding="2" width="350">
                <tr>
<td width="342" background=../image/tbg.gif><font color="FFFFFF">�������ʾ</font></td>
<td width="18">
<table border="1" bordercolorlight="666666" bordercolordark="FFFFFF" cellpadding="0" bgcolor="E0E0E0" cellspacing="0" width="18">
<tr>
<td width="16"><b><a href="javascript:history.go(-1)" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="�ر�"><font color="000000">��</font></a></b></td>
</tr>
</table>
</td>
</tr>
</table>
<table border="0" width="350" cellpadding="4">
<tr> 
                  <td width="59" align="center" valign="top"><font face="Wingdings" color="#0099FF" style="font-size:32pt">L</font></td>
<td width="269">
<%Select Case id
Case "000"%>
<p>�����Բ����㲻�ǹٸ�������Ա�����Ƭ�����ˣ���Ȩɾ��������Ƭ��</p>

<%Case else%>
<p>�����Բ��𣬸ó�������δ���Ǽǣ�����ϵվ�������</p>
<%End Select%>
</td>
</tr>
<tr>
<td colspan="2" align="center" valign="top">
<input type="button" name="ok" value="��ȷ ����" onclick=javascript:history.go(-1)>
</td>
</tr>
</table>
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
<script Language="JavaScript">
document.forms[0].ok.focus();
</script>
</body>
</html>

