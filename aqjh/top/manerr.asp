<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
id=Trim(Request.QueryString("id"))
%><html>
<head>
<title>出错提示</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/readonly/style.css">
</head>
<body background=../JHIMG/Bk_hc3w.gif leftmargin="0" topmargin="0">
<table border="0" width="400" cellpadding="0" cellspacing="0" align="center" background="../JHIMG/err_bg.gif">
  <tr> 
    <td width="98" align="center" valign="top"><img src="../Error.gif" width="76" height="76"></td>
    <td width="302"> 
      <%Select Case id
Case "255"%>
      <p>　　此项功能只有站长才可以查看！</p>
      <%Case else%>
      <p>　　对不起，该出错类型未被登记，请联系站长解决。</p>
      <%End Select%>
    </td>
  </tr>
  <tr> 
    <td colspan="2" align="center" valign="top" height="17"> <a href="javascript:history.go(-1)" 
            onMouseOut=MM_swapImgRestore() 
            onMouseOver="MM_swapImage('back','','../jhimg/err_but2.gif',1)"><img 
            border=0 height=20 name=back src="../jhimg/err_but1.gif" 
            width=73></a><br>
      <img height=5 src="../jhimg/err_bot.gif" 
        width=400></td>
  </tr>
</table>
</body>
</html>