<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<html>
<head>
<title><%=Application("aqjh_chatroomname")%>出题答对有奖</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../../css.css">
</head>
<%
%>
<form method="post" action="asking.asp">
<table width="250" border="1" cellspacing="0" cellpadding="3" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr align="center" bgcolor="#33CCCC">
<td colspan="2" height="27"><%=Application("aqjh_chatroomname")%>抢答系统出题</td>
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
<input type="text" name="reply" size="20" maxlength="15">
</td>
</tr>
<tr>
<td> 奖：</td>
<td>
<input type=text name='silver' value='100' maxlength="5" size=8>
</td>
</tr>
<tr>
<td colspan="2">
<div align="center">
<input type="submit" name="Submit" value="提 交">
<input type="button" name="button" onclick="window.close()" value="关 闭">
</div>
</td>
</tr>
<tr>
<td colspan="2">
<div align="center">出题人：<%=aqjh_name%> </div>
</td>
</tr>
</table>
</form>
<p align="center">网管负责出题，对于答对的聊友，系统将给予奖励！<br>
注意!不得利用此功能做假! </p>
</body>
</html>