<%
if Application("Ba_jxqy_ask")="" then 
response.redirect "../error.asp?id=061"
response.end
end if
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
%>
<html>
<head>
<title>答对有奖</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css1.css">
<script language=javascript>window.moveTo(100,50);window.resizeTo(screen.availWidth*2/3,screen.availHeight*3/4);</script>
</head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>'>
<form method="post" action="replying.asp">
  <table width="250" border="1" cellspacing="0" cellpadding="3" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr align="center" bgcolor="#33CCCC"> 
      <td colspan="2" height="27">神话虚幻抢答</td>
    </tr>
    <tr> 
      <td width="32" valign="top">问：</td>
      <td width="202"> 
        <textarea name="ask" readonly rows="3" wrap="VIRTUAL" cols="30"><%=Application("Ba_jxqy_ask")%></textarea>
      </td>
    </tr>
    <tr> 
      <td width="32">答：</td>
      <td width="202"> 
        <input type="text" name="reply" size="10" maxlength="10">
        <input type="submit" name="Submit" value="提 交">
        <input type="button"  value="关 闭" onclick='javascript:top.window.close();'>
      </td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <div align="center">答题人：<%=session("Ba_jxqy_username")%> </div>
      </td>
    </tr>
  </table>
</form>
</body>
</html>
