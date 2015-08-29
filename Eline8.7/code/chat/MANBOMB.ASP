<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
grade=Int(sjjh_grade)
inthechat=Session("sjjh_inthechat")
if grade<7 then Response.Redirect "../error.asp?id=482"
if inthechat<>"1" then Response.Redirect "manerr.asp?id=482"
bombname=Trim(Request.QueryString("id"))
if bombname="" then Response.Redirect "../error.asp?id=481"
if CStr(bombname)=CStr(sjjh_name) then Response.Redirect "../error.asp?id=481"%><html>
<head>
<title>炸弹操作</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="readonly/style.css">
<script language=javascript>window.moveTo(100,50);window.resizeTo(screen.availWidth*2/3,screen.availHeight*3/4);</script>
</head>
<body bgcolor="#FFFFFF" class=p150>
<div align="center">
<h1><font color="0099FF"><font size="6" face="隶书" color="#000000">炸 弹 操 作</font></font></h1>
</div>
<hr noshade size="1" color=009900 width="70%">
<div align="center">
<table width="380" border="0" cellspacing="2" cellpadding="2">
<tr>
<td><b>〔注意事项〕</b> <br>
必须输入扔炸弹的原因才能投放炸弹，被炸的对象的机器将不停地弹出新窗口，直到系统资源耗尽、死机。<br>
炸弹操作会被记录在“聊务公开”栏中，供广大聊友监督，因此，请勿随意炸人！ </td>
</tr>
</table>

</div>
<hr noshade size="1" color=009900 width="70%">
<br>
<table border="1" cellspacing="0" cellpadding="3" bgcolor="E0E0E0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" width="380">
<form method="post" action="manbombok.asp">
<tr>
<td>
<table border="0" cellpadding="4" width="100%" cellspacing="0">
<tr>
<td align="right" width="26%">轰炸用户名：</td>
<td width="74%"><font color="#FF0000"><%=bombname%>
<input type="hidden" name="bombname" value="<%=bombname%>">
</font></td>
</tr>
<tr>
<td align="right" width="26%">轰炸的原因：</td>
<td width="74%">(&lt;=60字符) </td>
</tr>
<tr>
<td align="right" width="26%">
<select name="select" onchange="javascript:document.forms[0].bombwhy.value=this.value;document.forms[0].bombwhy.focus();">
<option value="" selected>自填</option>
<option value="不受欢迎人。">不欢迎</option>
<option value="所取的名字十分不雅。">不雅ID</option>
<option value="乱刷屏，警告又不听。">乱刷屏</option>
<option value="在聊天室散布有悖伦理道德的言论。">不道德</option>
<option value="不遵守聊天规则，进行人身攻击。">骂人</option>
<option value="在聊天室散布违反国家法律法规的言论。">违法</option>
</select>
：</td>
<td width="74%">
<input type="text" name="bombwhy" maxlength="60" size="40">
</td>
</tr>
</table>
<div align="center">
<input type="submit" name="bombok" value="轰炸">
<input type="button" value="放弃" onclick="javascript:history.go(-1)">
</div>
</td>
</tr>
</form>
</table>
<br>
<hr noshade size="1" color=009900 width="70%">
<div align=center class=cp><%Response.Write "序列号：<font color=blue>" & Application("sjjh_sn") & "</font>，授权给：<font color=blue>" & Application("sjjh_user") & "</font><br><font color=999999>" & Application("sjjh_copyright") & "</font>"%></div>
</body>
</html>