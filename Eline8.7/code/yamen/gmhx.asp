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
%>
<html>
<head>
<title>〖<%=Application("sjjh_chatroomname")%>〗♀wWw.happyjh.com♀</title>
<style type="text/css">body, td     { font-size: 14 }
input        { font-size: 14; color: #000000 }
.p1          { font-size: 21pt; color: #ff0000 }
.p2          { font-size: 9pt; color: #00ee00 }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>

<body bgcolor="#006699">
<table border="1" bgcolor="#000000" align="center" width="360" cellpadding="10"
cellspacing="13" height="126">
<tr>
    <td bgcolor="#88AFD7"> 
      <form method="POST" action="gmhxok.asp" autocomplete="off">
        <table width="100%">
          <tr> 
            <td height="10"><font size="-1">注册ID:  
              <input type="password" name=id size=8 maxlength="5" onKeyDown="ctlent()">
              </font> </td>
            <td height="10"><font size="-1"> 原名:  
              <input type="text" name="name" size="10" maxlength="10" onKeyDown="ctlent()">
              </font></td>
          </tr>
          <tr> 
            <td height="2"><font size="-1">密码:  
              <input type="password" name="pass" size="10" maxlength="10" onKeyDown="ctlent()">
              </font> </td>
            <td height="2"> <font size="-1">新名:  
              <input type="text" name="name1" size="10" maxlength="5" onKeyDown="ctlent()">
              </font></td>
          </tr>
          <tr>
            <td align="center" colspan="2"><br> 
              <input type="submit" value="修改" name="submit">
              <input type="reset" value="取消" name="reset">
            </td>
          </tr>
          <tr> 
            <td align="center" colspan="2"><b><font color="blue" size="2">代价:</font><font color="#FF0000" size="2">现金5亿/金币200/会员降1级/丢失论坛信息！</font></b></td>
          </tr>
        </table>
</form>
</td>
</tr>
</table>
</body>
</html>