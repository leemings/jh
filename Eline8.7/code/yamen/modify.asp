<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
randomize timer
regjm=int(rnd*9998)+1
%>
<html>
<head>
<title>修改密码♀wWw.51eline.com♀</title>
<LINK href="../css.css" rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#006699">
<center>
  <table width=360 border=1 align=center cellpadding="5" cellspacing="10" bgcolor="#000000">
    <tr bgcolor="#FFFFFF" align="center">
      <td height="209" bgcolor="#006699"> 
        <p><font color="#FFFFFF">修 改 密 码</font></p>
        <table border="0" width="288" height="257">
          <tr>
<td>
<form method=POST action='modifyok.asp'>
                <div align="center"><font color="#FFFFFF">江湖ID：</font> 
                  <input type=password name=id size=12 maxlength="10">
                  <br>
                  <font color="#FFFFFF">姓　名：</font> 
                  <input type=text name=name size=12 maxlength="10">
                  <br>
                  <font color="#FFFFFF">原密码：</font> 
                  <input type=password name=oldpass size=12 maxlength="10">
                  <br>
                  <font color="#FFFFFF">新密码：</font> 
                  <input type=password name=pass size=12 maxlength="10">
                  <br>
                  <font color="#FFFFFF">确　认：</font> 
                  <input type=password name=repass size=12 maxlength="10">
                  <br>
                  <br>
                  <br>
                  <p> 
                    <input type=submit value=修改 name="submit">
                    <input type=button value=关闭 onClick="window.close()" name="button">
                  </p>
                </div>
</form>
</td>
</tr>
<tr>
<td style='color:red;font-size:9pt'> <font color="#FF6600">说明：<br>
由于IE5会记录输入的密码，<br>
请在网吧上网的朋友通过“清除历史记录”<br>
来删除记录！以免帐号被盗用 </font></td>
</tr>
</table>
</table>

</center>
</body>
</html>
