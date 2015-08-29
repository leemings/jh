<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"

%>
<html>
<head>
<title><%=Application("sjjh_chatroomname")%>出题答对有奖</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css1.css">
<script language=javascript>window.moveTo(100,50);window.resizeTo(screen.availWidth*2/3,screen.availHeight*3/4);</script>
</head>
<%
%>
<form method="post" action="asking.asp">
<table width="250" border="1" cellspacing="0" cellpadding="3" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr align="center" bgcolor="#33CCCC">
<td colspan="2" height="27"><%=Application("sjjh_chatroomname")%>抢答系统出题</td>
</tr>
<tr>
<td width="32" valign="top">问：</td>
<td width="202">
<textarea name="ask" rows="3" wrap="VIRTUAL" cols="30"></textarea>
</td>
</tr>
<tr>
<td width="32">答：</td>
<td width="202">
<input type="text" name="reply" size="15" maxlength="15">
</td>
</tr>
<tr>
<td> 奖：</td>
<td>
<input type=text name='silver' value='100' maxlength=5 size=5>
</td>
</tr>
<tr>
<td colspan="2">
<div align="center">
<input type="submit" name="Submit" value="提 交">
<input type="submit" name="Submit2" value="关 闭">
</div>
</td>
</tr>
<tr>
<td colspan="2">
<div align="center">出题人：<%=sjjh_name%> </div>
</td>
</tr>
</table>
</form>
<p align="center">网管负责出题，对于答对的聊友，系统将给予奖励！<br>
注意!不得利用此功能做假! </p>
</body>
</html>
