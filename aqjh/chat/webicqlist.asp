<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
useronlinename=Application("aqjh_useronlinename"&session("nowinroom"))
if aqjh_name="" or Session("aqjh_inthechat")<>"1" or Instr(useronlinename," "&aqjh_name&" ")=0 then Response.Redirect "chaterr.asp?id=001"
chatroombgimage=Application("aqjh_chatimage")
chatroombgcolor=Application("aqjh_chatbgcolor")%>
<html>
<head>
<title>��Ϣ�б�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
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
<td width=326><font color=FFFFFF face=Wingdings>*</font><font color=FFFFFF>���� Web ICQ ��Ϣ</font></td>
<td width=18>
<table border=1 bordercolorlight=666666 bordercolordark=FFFFFF cellpadding=0 bgcolor=E0E0E0 cellspacing=0 width=18>
<tr>
<td width=16><b><a href="javascript:window.close()" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="�˳�"><font color="000000">��</font></a></b></td>
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
<td width="25%" bgcolor="E0E0E0"><a href=webicq.asp><font color="#000000">������Ϣ</font></a></td>
<td width="25%" bgcolor="#000000"><font color="#FFFFFF">��Ϣ�б�</font></td>
<td width="25%">&nbsp;</td>
<td width="25%">&nbsp;</td>
</tr>
</table>
<table width="100%" border="1" height="200" cellspacing="0" cellpadding="5" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="E0E0E0">
<form>
<tr>
<td>
<div align="center">������Ϣδ����ȡ��ֻ����10����<br>
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
