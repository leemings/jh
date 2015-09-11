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
if Application("aqjh_ask")="" then
response.redirect "../../error.asp?id=476"
response.end
end if
username=aqjh_name
%>
<html>
<head>
<title><%=Application("aqjh_chatroomname")%>答对有奖</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../../css.css">
<script language=javascript>window.moveTo(100,50);window.resizeTo(screen.availWidth*2/3,screen.availHeight*3/4);</script>
</head>
<%
%>
<form method="post" action="replying.asp">
<table width="300" border="1" cellspacing="0" cellpadding="3" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr align="center" bgcolor="#33CCCC">
<td colspan="2" height="27"><%=Application("aqjh_chatroomname")%>抢答器</td>
</tr>
<tr>
<td width="63" valign="top">问：</td>
<td width="214">
<textarea name="ask" readonly rows="3" wrap="VIRTUAL" cols="30"><%=Application("aqjh_ask")%></textarea>
</td>
</tr>
<tr>
<td width="63">答：</td>
<td width="214">
<input type="text" name="reply" size="20" maxlength="50">
</td>
</tr>
<tr>
<td width="63">奖励：</td>
<td width="214"> <%=Application("aqjh_asksilver")%>两</td>
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
<div align="center">答题人：<%=aqjh_name%> </div>
</td>
</tr>
</table>
</form>
</body></html>