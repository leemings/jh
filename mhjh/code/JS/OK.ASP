
<%@ LANGUAGE=VBScript%> 
<%Response.Expires=0 
id=Trim(Request.QueryString("id")) 
Select Case id 
Case "100" 
nl="交易成功。" 
End Select 
nl="<p>　　" & nl & "</p>"%><html> 
<head> 
<title>操作成功</title> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"> 
<link rel="stylesheet" href="../style.css"> 
</head> 
<body oncontextmenu=self.event.returnValue=false  leftmargin="0" topmargin="0" background='../chatroom/bg1.gif'> 
<table width="100%" border="0" height="415"> 
<tr align="center"> 
<td height="411"> 
<form method="post" action=""> 
<table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" bgcolor="E0E0E0"> 
<tr> 
<td> 
<table border="0" bgcolor="#009900" cellspacing="0" cellpadding="2" width="350"> 
<tr> 
<td width="342"><font color="FFFFFF" face="Wingdings">z</font><font color="FFFFFF">操作成功</font></td> 
<td width="18"> 

</td> 
</tr> 
</table> 
<table border="0" width="350" cellpadding="4"> 
<tr> 
<td width="269"> 
<%=nl%> <br>
<p align=center><a href="bank.asp" onmouseover="window.status='返回钱庄';return true;" onmouseout="window.status='';return true;"><font color="#000000">返回</font></a></p></body> 

</form> 
<script Language="javasctipt "> 
document.forms[0].ok.focus(); 
</script> 
