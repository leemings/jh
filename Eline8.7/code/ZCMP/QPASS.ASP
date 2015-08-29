<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
session("sjjh_jm")=session("sjjh_jm")+1
if session("sjjh_jm")>30 then Response.Redirect "../chat/readonly/bomb.htm"
randomize timer
regjm=int(rnd*9998)+1
%>
<html>
<head>
<title>〖<%=Application("sjjh_chatroomname")%>〗♀wWw.51eline.com♀</title>
<style type="text/css">body, td     { font-size: 14 }
input        { font-size: 14; color: #000000 }
.p1          { font-size: 21pt; color: #ff0000 }
.p2          { font-size: 9pt; color: #00ee00 }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>

<body bgcolor="#006699" leftmargin="0" topmargin="3">
<p class="p1" align="center"><b><font color="#FFFFFF" size="3">取回密码</font></b></p>
<table border="1" bgcolor="#000000" align="center" width="327" cellpadding="10"
cellspacing="13">
  <tr>
    <td bgcolor="#FFFFFF"> 
      <form method="POST" action="qpassok.asp">
        <table width="100%">
          <tr> 
            <td height="10"> 
              <div align="center"><font size="-1">注册ID:  
                <input type=text name=id size=8 maxlength="5">
                </font> </div>
            </td>
          </tr>
          <td height="10"> 
            <div align="center"><font size="-1"> 提示问题:  
              <input type="text" name="wt" size="10" maxlength="10">
              </font></div>
          </td>
          <tr> 
            <td height="2"> 
              <div align="center"><font size="-1">问题答案</font><font size="-1">:  
                <input type="text" name="da" size="10" maxlength="10">
                </font> </div>
            </td>
            <td height="2"> <font size="-1">&nbsp; </font></td>
          </tr>
          <tr> 
            <td align="center" colspan="2"> 
              <input type="submit" value="确定"
name="submit">
              <input type="reset" value="取消" name="reset">
            </td>
          </tr>
          <tr> 
            <td align="center" colspan="2"><b><font color="#FF0000" size="2">请输入ID、密码提示问题、答案</font></b></td>
          </tr>
        </table>
</form>
</td>
</tr>
</table>

<p align="center"><font color="#FFFFFF" size="2">『E线江湖』</font></p>

</body>

</html>