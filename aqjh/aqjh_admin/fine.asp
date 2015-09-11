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
<title>用户数据库管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css/css.css" type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center"><%=Application("aqjh_chatroomname")%>数据库</p>
<p align="left">管理说明：<br>
  <font color="#0000FF">江湖战斗等级计算公式：等级x等级x50，如：开一个战斗10级的用户：10x10x50=5000将这个值写入到数据库的：<b>[<font color="#FF0000">总积分</font></b>]里，否则在用户下一次进入时战斗等级会自动降下来。</font><font color="#FF0000"><br>
  </font><font color="#0000FF">会员时间是截止时间，会员等级为：1,2,3,4如果其值为0，则表示不是会员！</font><br>
  <font color="#FF0000"> 武功加，内力加，体力加，攻击加等。这一些值是不要乱改的，这些是他们在江湖上打坐、闭目得到的这些值不要改动！<br>
  <font color="#000000">关于泡分制会员说明：</font> <br>
  [<b><font color="#0000FF">会员</font></b>]：当为True 时则为泡分制会员，[<b><font color="#0000FF">会员结束</font></b>]则是泡分制会员的结束时间。有关泡分制会员默认为2倍，要改修改可以修改泡分文件得到！</font></p>
<table width=80% align=center><tr><td>
<form method="POST" action="showuser.asp">
    <select name="sjcz">
      <option value="姓名" selected>姓名</option>
      <option value="ID">ID</option>
      <option value="配偶">配偶</option>
      <option value="信箱">信箱</option>
      <option value="Oicq">OIcq</option>
    </select>
    <input type="text" name="search" size="10" maxlength="10">
    <input type="submit" value="查询" name="B1" class="p9">
    <input type="reset" name="Submit" value="清空"> [用户资料查询]
</form>
<form method="POST" action="setzt.asp">
    <select name="t1">
      <option value="银两" selected>银两</option>
      <option value="存款">存款</option>
      <option value="金币">金币</option>
      <option value="会员金卡">金卡</option>
      <option value="道德">道德</option>
      <option value="魅力">魅力</option>
      <option value="内力">内力</option>
      <option value="体力">体力</option>
      <option value="武功">武功</option>
      <option value="攻击">攻击</option>
      <option value="防御">防御</option>
      <option value="转生">转生</option>
      <option value="花魁">花魁</option>
      <option value="法力">法力</option>
      <option value="轻功">轻功</option>
    </select>
    姓名：<input type="text" name="name" size="10" maxlength="10">
    <select name="t2">
      <option value="增加" selected>增加</option>
      <option value="减少">减少</option> 
</select>
<input type="text" name="sl" size="10" maxlength="9">
	<input type="submit" value="确定" name="B1" class="p9">
    <input type="reset" name="Submit" value="清空"> [快速状态调整]
</form>
<form method="POST" action="laren.asp">
    <input type="text" name="larenseek" size="10" maxlength="10">
    <input type="submit" value="查询" name="B12" class="p9">
    <input type="reset" name="Submit2" value="清空"> [拉人查看程序]
</form>
<form method="POST" action="chuwu.asp">
    <input type="text" name="chuwuseek" size="10" maxlength="10">
    <input type="submit" value="查询" name="B12" class="p9">
    <input type="reset" name="Submit2" value="清空"> [查看储物箱，按ID查询]
</form>
<form method="POST" action="pass.asp">
    <input type="text" name="cpass" size="10" maxlength="10">
    <input type="submit" value="修改" name="B12" class="p9">
    <input type="reset" name="Submit2" value="清空"> [输入人名将登录密码改成：123456]
</form>
<form method="POST" action="pass2.asp">
    <input type="text" name="cpass" size="10" maxlength="10">
    <input type="submit" value="修改" name="B12" class="p9">
    <input type="reset" name="Submit2" value="清空"> [输入人名将第二密码改成：123456]
</form>
<form method="POST" action="yuejlok.asp">
    <input type="text" name="jl_name" size="10" maxlength="10">
    <input type="submit" value="确定" name="B12" class="p9">
    <input type="reset" name="Submit2" value="清空"> [本月奖励(包括月积分和拉人排行)]
</form></td></tr></table></body></html>