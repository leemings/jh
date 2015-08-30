<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%><html>
<head>
<title>睡眠解救♀wWw.happyjh.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../css.css" rel=stylesheet>
</head>
<body leftmargin="0" topmargin="0" bgcolor="#006699">
<table width="360" border="1" align="center" cellpadding="8"
cellspacing="10" bgcolor="#000000">
  <tr bgcolor="#FFFFFF" align="center">
    <td bgcolor="#006699"> <font color="#FFFFFF">请大侠仔细填写以下信息</font><font color="#CCCCCC">------</font><font color="#FF3333">睡眠解救</font> 
      <form method="POST" action="jiejiu4.asp">
        <table width="254" align="center" height="194">
          <tr>
            <td height="126"> 
              <div align="center"> <font color="#FFFFFF">注册ID：</font> 
                <input type=password name=id size=8 maxlength="10">
                <br>
                <br>
                <font color="#FFFFFF">姓名： </font>
<input type="text" name="name" size="10" maxlength="10">
                <br>
                <br>
                <font color="#FFFFFF">密码：</font> 
                <input type="password" name="pass" size="10" maxlength="10">
                <p>
              </div>
</td>
</tr>
<tr>
<td align="center">
<input type="submit" name="submit" value="解救">
<input type="button" value="关闭" onclick="window.close()"
name="button"></td>
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