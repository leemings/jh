<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
useronlinename=Application("sjjh_useronlinename"&session("nowinroom"))
if sjjh_name="" or Session("sjjh_inthechat")<>"1" or Instr(useronlinename," "&sjjh_name&" ")=0 then Response.Redirect "chaterr.asp?id=001"
chatroombgimage=Application("sjjh_chatimage")
chatroombgcolor=Application("sjjh_chatbgcolor")%>
<html>
<head>
<title>消息列表</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor=<%=chatroombgcolor%> background=<%=chatroombgimage%> leftmargin="0" topmargin="0">
<table width="100%" border="0" height="100%">
<tr>
<td>
<table width=350 border=1 align=center cellspacing=0 bordercolorlight=000000 bordercolordark=FFFFFF bgcolor=E0E0E0>
<tr valign="top">
<td>
<table border=0 bgcolor=#3399FF cellspacing=0 cellpadding=2 width=344>
<tr>
<td width=326><font color=FFFFFF face=Wingdings>*</font><font color=FFFFFF>发送 Web ICQ 消息</font></td>
<td width=18>
<table border=1 bordercolorlight=666666 bordercolordark=FFFFFF cellpadding=0 bgcolor=E0E0E0 cellspacing=0 width=18>
<tr>
<td width=16><b><a href="javascript:window.close()" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="退出"><font color="000000">×</font></a></b></td>
</tr>
</table>
</td>
</tr>
</table>
<table width="100%" border="1" cellspacing="0" bordercolorlight="#999999" bordercolordark="#FFFFFF" bgcolor="FFFFFF" cellpadding="0">
<tr valign="middle" align="center">
<td class=p9>
<table width="100%" border="0" cellpadding="3">
<tr align="center">
<td width="25%" bgcolor="E0E0E0"><a href=webicq.asp><font color="#000000">发送消息</font></a></td>
<td width="25%" bgcolor="#000000"><font color="#FFFFFF">消息列表</font></td>
<td width="25%">&nbsp;</td>
<td width="25%">&nbsp;</td>
</tr>
</table>
<table width="100%" border="1" height="200" cellspacing="0" cellpadding="5" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="E0E0E0">
<form>
<tr>
<td>
<div align="center">下列消息未被领取，只保留10分钟<br>
<iframe name="gb" frameborder="0" scrolling="auto" src="webicqlistok.asp" height="170" width="320"></iframe>
</div>
</td>
</tr>
</form>
</table>
</td>
</tr>
</table>
</td>
</tr>
</table>
</td>
</tr>
</table>
</body>
</html>
