<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%><html>
<head>
<title>◎ 阎王殿门口-你碰见了观音大士</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../css.css" rel=stylesheet>
</head>
<body leftmargin="0" topmargin="0" bgcolor="#006699">
<table width="360" border="1" align="center" cellpadding="8"
cellspacing="10" bgcolor="#000000">
  <tr bgcolor="#FFFFFF" align="center">
    <td bgcolor="#006699"> <font color="#FFFFFF">请大侠仔细填写以下信息</font><font color="#CCCCCC">------</font><font color="#FF3333">报仇雪恨</font> 
      <form method="POST" action="casper.asp">
        <table width="254" align="center" height="194">
          <tr>
            <td height="126"> 
              <div align="center">
                <font color="#FFFFFF">姓名：</font> 
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
<input type="submit" name="submit" value="复活">
<input type="button" value="关闭" onclick="window.close()"
name="button"></td>
</tr>
<tr>
            <td><font color="#FFFFFF"><strong>阎王说:</strong></font><br>
              <font color="#FFFFFF">一定要记住这个日子!!!这是你死去的日子,我可以让你重生,你一定要复仇!!!</font><br>
</table>
</form>
</table>

</body>

</html>