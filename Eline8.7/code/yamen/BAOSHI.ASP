<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%><html>
<head>
<title>自我保释♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../css.css" rel=stylesheet>
</head>
<body leftmargin="0" topmargin="0" bgcolor="#006699">
<table width="360" border="1" align="center" cellpadding="8"
cellspacing="10" bgcolor="#000000">
  <tr bgcolor="#FFFFFF" align="center">
    <td bgcolor="#006699"> <font color="#FFFFFF">请大侠选择你要保释自己的方式</font><font color="#CCCCCC">------</font><font color="#FF3333">保释出牢</font> 
      <form method="POST" action="baoshi1.asp">
        <table width="254" align="center" height="194">
          <tr>
            <td height="67"> 
              <p align="center"><a href="baoshi1.ASP"><b><font color="#FFFF00" size="5">现金保释</font></b></a><b></b>
            </td>
</tr>
          <tr>
            <td height="61"> 
              <p align="center"><a href="baoshi2.ASP"><b><font color="#FFFF00" size="5">存款保释</font></b></a><b></b>
            </td>
          </tr>
<tr>
<td>
  <p align="center"><font color="#FF0000">此功能由<%=Application("sjjh_chatroomname")%>首创</font><br>
                <font color="#FFFFFF">一定有要1亿银两/20金币才可以，等着坐牢房过日子吧！</font><br>
                <font color="#FF0000">注意：本功能只对逮捕有用，被坐牢和监禁者无法保释。</font> </table>
</form>
</table>

</body>

</html>