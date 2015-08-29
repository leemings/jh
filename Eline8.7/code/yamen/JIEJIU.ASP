<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%><html>
<head>
<title>睡眠解救♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../css.css" rel=stylesheet>
</head>
<body leftmargin="0" topmargin="0" bgcolor="#006699">
<table width="360" border="1" align="center" cellpadding="8"
cellspacing="10" bgcolor="#000000">
  <tr bgcolor="#FFFFFF" align="center">
    <td bgcolor="#006699"> <font color="#FFFFFF">请大侠选择你解救自己的方式</font><font color="#CCCCCC">------</font><font color="#FF3333">睡眠解救</font> 
      <form method="POST" action="jiejiu1.asp">
        <table width="254" align="center" height="194">
          <tr>
            <td height="64"> 
              <p align="center"><a href="jiejiu1.asp"><font color="#FFFF00" size="5"><strong>现金解救</strong></font></a> 
            </td>
</tr>
          <tr>
            <td height="64"> 
              <p align="center"><a href="jiejiu2.asp"><font color="#FFFF00" size="5"><strong>存款解救</strong></font></a> 
            </td>
          </tr>
<tr>
<td>
  <p align="center"><font color="#FF0000">此功能由<%=Application("sjjh_chatroomname")%>首创</font><br>
                <font color="#FFFFFF">睡眠解救自己一定要有[5000万银两/5000点法力/20金币]才行，没有的就乖乖的去睡大觉吧！</font> 
            </table>
</form>
</table>

</body>

</html>