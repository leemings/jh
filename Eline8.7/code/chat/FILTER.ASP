<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
filname=Session("sjjh_filname")
sjjh_name=sjjh_name
useronlinename=Application("sjjh_useronlinename"&session("nowinroom"))
if sjjh_name="" or Session("sjjh_inthechat")<>"1" or Instr(useronlinename," "&sjjh_name&" ")=0 then Response.Redirect "chaterr.asp?id=001"
show=Split(Trim(useronlinename)," ",-1)
j=UBound(show)
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"%><html>
<head>
<META http-equiv='content-type' content='text/html; charset=gb2312'>
<title>屏蔽某人的话♀一线网络→wWw.happyjh.com♀</title>
<style type='text/css'>
body {font-size:12pt;}
td {font-size:10.5pt;line-height:130%}
input{font-size:9pt}
select{font-size:9pt}
textarea{font-size:9pt}
.p9{font-size:9pt}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed">
<div align="center">
<table width="100%" border="0" height="100%">
<tr>
<td>
<table border="1" cellspacing="0" bgcolor="E0E0E0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" width="100%" cellpadding="3">
<form method="post" action="filterok.asp" name="">
<tr align="center">
<td><font color="red">选择屏蔽对像</font>
<table width="100%" border="0" cellspacing="0" cellpadding="4">
<tr>
<td><%for i=0 to j
if CStr(show(i))<>CStr(sjjh_name) then%><input type="checkbox" name="filtername" value="<%=show(i)%>"<%if Instr(filname," "&show(i)&",")<>0 then Response.Write " checked"%>><%=show(i)%><br>
<%end if
next%></td>
</tr>
</table>
<table width="100%" border="0">
<tr align="center">
<td>
<input type="submit" name="ok" value="提交">
<input type="submit" name="clearall" value="取消">
</td>
</tr>
</table>
</td>
</tr>
</form>
</table>
</td>
</tr>
</table>
</div>
<Script Language=Javascript>
parent.m.location.href='about:blank';
</Script>
</body></html>