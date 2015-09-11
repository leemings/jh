<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<!--#include file="config.asp"-->
<html>
<head>
<title>会员数据库管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css/css.css" type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center"><%=Application("aqjh_chatroomname")%>数据库</p>
<table width="572" align="center" border="1" align="center" cellspacing="1" cellpadding="4" bordercolor="#B8AF86">
<tr>
<td>
<p align="center">&nbsp;</p>
<form method="POST" action="kfaddadminok.asp">
<div align="center">
会员等级：
<select name="hygrade">
<option value="1" selected>一级会员</option>
<option value="2">二级会员</option>
<option value="3">三级会员</option>
<option value="4">四级会员</option>
<option value="5">五级会员</option>
<option value="6">六级会员</option>
<option value="7">七级会员</option>
<option value="8">八级会员</option>
</select>
<br>
金卡数目：
<input type="text" name="kfnum" size="10" maxlength="10">
<br>
<input type="submit" value="开始批量送卡" name="B1" class="p9">
</div>
<br>
</div>
<div align="center">
<center>
</center>
</div>
</form>
<p align="center">[站长可以通过此处来修改用户的所以资料]</p>
<p align="center"><a href="../jhmp/hy.asp">现有会员列表</a> <a href="fine.asp">更新用户资料</a></p>
</td>
</tr>
</table>
<p align="center">&nbsp;</p>
</body>
</html>
