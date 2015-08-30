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
if Application("sjjh_ask")="" then
response.redirect "../error.asp?id=476"
response.end
end if
username=sjjh_name
%>
<html>
<head>
<title><%=Application("sjjh_chatroomname")%>答对有奖♀wWw.happyjh.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css1.css">
<script language=javascript>window.moveTo(100,50);window.resizeTo(screen.availWidth*2/3,screen.availHeight*3/4);</script>
</head>
<%
%>
<form method="post" action="replying.asp">
<table width="285" border="1" cellspacing="0" cellpadding="3" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr align="center" bgcolor="#33CCCC">
<td colspan="2" height="27"><%=Application("sjjh_chatroomname")%>抢答器</td>
</tr>
<tr>
<td width="53" valign="top">问：</td>
<td width="214">
<textarea name="ask" readonly rows="3" wrap="VIRTUAL" cols="30"><%=Application("sjjh_ask")%></textarea>
</td>
</tr>
<tr>
<td width="53">答：</td>
<td width="214">
<input type="text" name="reply" size="10" maxlength="50">
</td>
</tr>
<tr>
<td width="53">奖励：</td>
<td width="214"> <%=Application("sjjh_asksilver")%>两</td>
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
<div align="center">答题人：<%=sjjh_name%> </div>
</td>
</tr>
</table>
</form>
</body>
</html>
