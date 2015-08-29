<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
session("sjjh_jm")=session("sjjh_jm")+1
if session("sjjh_jm")>30 then Response.Redirect "../chat/readonly/bomb.htm"
randomize timer
regjm=int(rnd*9998)+1
%><html>
<head>
<title>我不小心掉线</title>
<LINK href="../css.css" rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#006699">
<table border="1" bgcolor="#000000" align="center" width="360" cellpadding="10"
cellspacing="13" height="126">
  <tr bgcolor="#FFFFFF">
    <td height="75" bgcolor="#006699"> 
      <form method="POST" action="getidok.asp">
        <table width="250" align="center">
          <tr> 
            <td>
              <div align="center"><br>
                <font color="#FFFFFF">姓名：</font> 
                <input type="text" name="name" size="10" maxlength="10">
                <br>
                <font color="#FFFFFF">密码：</font> 
                <input type="password" name="pass" size="10" maxlength="10">
                <br>
                <font color="#FFFFFF">认证：</font> 
                <input type=text name=regjm1 size=5 maxlength="5">
                <br>
                <font color="#FFFFFF">认证号：</font><font color="#FF0000"><%=regjm%></font> 
                <br>
                <p> 
              </div>
            </td>
          </tr>
          <tr> 
            <td align="center"> 
              <input type=hidden name=regjm value="<%=regjm%>">
              <input type="submit" name="submit" value="确定">
              <input type="button" value="关闭" onclick="window.close()"
name="button">
            </td>
          </tr>
        </table>
</form>
</table>
</body>
</html>
