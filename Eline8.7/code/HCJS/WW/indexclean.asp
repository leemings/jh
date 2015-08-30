<%@ LANGUAGE=VBScript codepage ="936" %>
<%

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"


%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_name="" then%>
<script language=vbscript>
MsgBox "对不起，你还没有登录江湖！"
window.close()
</script>
<%else
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=Application("sjjh_chatroomname")%>—温泉浴中心♀wWw.happyjh.com♀</title>
<link rel="StyleSheet" href="../../css.css" title="Contemporary">
</head>

<body leftmargin="0" topmargin="0">
<div align="center">
<center>
<table width="778" border="0" cellspacing="0" cellpadding="0">
<tr bgcolor="#FFFFFF">
<td width="86"></td>
<td width="33"><img src="../../images/yc4.gif" width="84" height="193"></td>
        <td width="140" valign="bottom">&nbsp;</td>
<td width="519" valign="bottom">&nbsp;</td>
</tr>
<tr bgcolor="#000000">
<td width="86"></td>
<td width="33"><img src="../../images/yc3.jpg" width="84" height="242"></td>
        <td width="140" valign="top">&nbsp;</td>
<td width="519" valign="top">
<form method="POST" action="CHECKSEX.ASP">
<br>
<br>
<table width="100%" cellspacing="1">
<tr>
                <td class="p2" width="100%"> <font color="#CCCCCC">欢 迎 进 入 E 线 
                  江 湖 龙 泉 洗 浴 中 心&nbsp;</font> </td>
</tr>
<tr >
                <td class="p3" width="100%" height="29"> <font color="#FFFFFF">进入洗浴后就不能领工资！！</font>体力1000 
                </td>
</tr>
<tr>
<td class="p2" width="100%" height="50">

<input type="radio" value="男"
name="SEX" checked>
<font color="#FFFFFF"> <font color="#CCCCCC">男宾部&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" value="女" name="SEX">
女宾部</font></font> </td>
</tr>
<tr>
<td class="p3" width="100%">


<p>&nbsp;</p>
<p>
<input type="submit" value="进入洗浴"
name="B1">
</p>
</td>
</tr>
</table>
</form>
<%=Application("sjjh_chatroomname")%>版权所有 阿男工作室修改</td>
</tr>
</table>
</center>
</div>

</body>

</html>
<%
end if
%>

<html><script language="JavaScript">                                                                  </script></html>


<html><script language="JavaScript">                                                                  </script></html>


<html></html>

<html></html>
<html></html>
<html></html>
<html></html>
<html></html>
<html></html>


