<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
useronlinename=Application("sjjh_useronlinename"&session("nowinroom"))

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if sjjh_name="" or Session("sjjh_inthechat")<>"1" or Instr(useronlinename," "&sjjh_name&" ")=0 then Response.Redirect "chaterr.asp?id=001"
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"%><html>
<head>
<title>�������</title>
<META http-equiv='content-type' content='text/html; charset=gb2312'>
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
<div align=center><font color="#FF0000" style="font-size:12pt">�������</font></div>
<p>���Ѿ����ֺ�����Ϊ <font color="#FF0000"><script>document.write(parent.f2.document.af.fs.value);</script>��</font>���о�����Ϊ <font color="#FF0000"><script>document.write(parent.f2.document.af.lh.value);</script>%</font> ��������������������ʹ���ֺ���Ч��</p>
<div align=center>
<input type="button" value="����" onclick=javascript:history.go(-1)>
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