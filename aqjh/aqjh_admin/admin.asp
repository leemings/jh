<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>
<title>超级管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<frameset cols="180,*" rows="*">
<frame src="menu.asp" noresize scrolling="AUTO" frameborder="NO" name="menu">
<frame src="txt.asp" noresize scrolling="AUTO" frameborder="NO" name="txt">
</frameset>
<noframes><body bgcolor="#FFFFFF">
</body></noframes>
</html>