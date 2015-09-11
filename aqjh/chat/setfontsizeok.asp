<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
useronlinename=Application("aqjh_useronlinename"&session("nowinroom"))

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if aqjh_name="" or Session("aqjh_inthechat")<>"1" or Instr(useronlinename," "&aqjh_name&" ")=0 then Response.Redirect "chaterr.asp?id=001"
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"%><html>
<head>
<title>设置完成</title>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed">
<div align=center>
<table width="100%" border="0" height="100%">
<tr>
<td>
<table border="1" bgcolor="E0E0E0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" cellpadding="4">
<form>
<tr>
<td style="font-size:10.5pt">
<div align=center><font color="#FF0000" style="font-size:12pt">设置完成</font></div>
<p>　已经将字号设置为 <font color="#FF0000"><script>document.write(parent.f2.document.af.fs.value);</script>磅</font>，行距设置为 <font color="#FF0000"><script>document.write(parent.f2.document.af.lh.value);</script>%</font> ，必须点击“清屏”才能使该字号生效。</p>
<div align=center>
<input type="button" value="返回" onclick=javascript:history.go(-1)>
</div>
</td>
</tr>
</form>
</table>
</td>
</tr>
</table>
</div>
</body>
</html>