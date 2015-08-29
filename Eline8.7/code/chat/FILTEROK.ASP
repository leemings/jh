<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
sjjh_name=sjjh_name
if sjjh_name="" or Session("sjjh_inthechat")<>"1" then Response.Redirect "close.asp"
filname=Request.Form("filtername")
filname=" "&filname&","
clearall=Request.Form("clearall")
if clearall="取消" then filname=" ,"
Session("sjjh_filname")=filname
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"%><html>
<head>
<title>屏蔽讨厌鬼</title>
<style type='text/css'>
body {font-size:12pt;}
td {font-size:10.5pt;line-height:120%}
input{font-size:9pt}
select{font-size:9pt}
textarea{font-size:9pt}
.p9{font-size:9pt}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed">
<table width="100%" border="0" height="100%">
<tr>
<td>
<table width="100%" border="1" cellspacing="0" cellpadding="4" bgcolor="E0E0E0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<form>
<tr align="center">
<td>
<font color="red">屏蔽结果</font><br>
<%if filname=" ," then%><p>关闭此功能</p><%else%>
<font class="p9">下列用户的发言将被屏蔽，不会再出现在你的屏幕上。</font>
<p><font color=blue><%=filname%></font></p><%end if%>
<input type="button" value="返回" onclick="javascript:history.go(-1)">
</td>
</tr>
</form>
</table>
</td>
</tr>
</table>
</body>
</html>