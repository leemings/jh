<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>
<title>『E线江湖』--超级管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<frameset cols="90,652" rows="*">
<frame src="menu.asp" noresize scrolling="yes" frameborder="NO" name="menu">
<frame src="txt.asp" noresize scrolling="AUTO" frameborder="NO" name="txt">
</frameset>
<noframes><body bgcolor="#FFFFFF">
</body></noframes>
</html>
