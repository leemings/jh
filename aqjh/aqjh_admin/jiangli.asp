<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
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
<title>奖励系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 宋体}
input {  font-size: 9pt; color: #000000; background-color: #f7f7f7; padding-top: 3px}
.c {  font-family: 宋体; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: normal; font-variant: normal; text-decoration: none}
--></style>
</head>

<body >
<p>&nbsp;</p>
<p align="center"><b><font size="3">快乐</font></b><b><font size="3">江湖奖励系统</font></b></p>
<form action="jiangli2.asp" method=POST >
  <div align="center"> 奖励  
    <select name="dj">
      <option value="会员等级<1" selected>非会员</option>
      <option value="会员等级>0">会员</option>
      <option value="grade=5">掌门</option>
      <option value="grade>6">官府</option>
    </select>
    <input type="text" name="sl" maxlength="2" size="5" value="10"> 
    名 ,  
    <input type="text" name="number" value="100" maxlength="5" size="10">
    <select name="xm"> 
      <option value="allvalue" selected>泡点</option>
      <option value="金币">金币</option>
      <option value="泡豆点数">豆点</option>
      <option value="会员金卡">金卡</option>
      <option value="银两">银两</option>
      <option value="道德">道德</option>
      <option value="魅力">魅力</option>
      <option value="内力">内力</option>
      <option value="体力">体力</option>
      <option value="武功">武功</option>
      <option value="攻击">攻击</option>
      <option value="防御">防御</option>
      <option value="存款">存款</option>
      <option value="法力">法力</option>
      <option value="轻功">轻功</option>
    </select>
    <input type="submit" name="Submit" value="奖励"></div> 
</form>
</body>
</html>
