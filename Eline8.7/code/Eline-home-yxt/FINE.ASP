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
<title>用户数据库管理♀wWw.happyjh.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>
<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif" topmargin="0">
<p align="center"><%=Application("sjjh_chatroomname")%>数据库</p>
<p align="left">管理说明：<br>
  <font color="#0000FF">江湖战斗等级计算公式：等级x等级x50，如：开一个战斗10级的用户：10x10x50=5000将这个值写入到数据库的：<b>[<font color="#FF0000">总积分</font></b>]里，否则在用户下一次进入时战斗等级会自动降下来。</font><font color="#FF0000"><br>
  </font><font color="#0000FF">会员时间是截止时间，会员等级为：1,2,3,4如果其值为0，则表示不是会员！</font><br>
  <font color="#FF0000"> 武功加，内力加，体力加，攻击加等。这一些值是不要乱改的，这些是他们在江湖上打坐、闭目得到的这些值不要改动！如果在管理上有什么不明白的问题，请找小小宇 
  QQ:81135 或到 <a href="http://xajh.xhong.net" target="_blank">江湖论谈留言</a><br>
  <font color="#000000">关于泡分制会员说明：</font> <br>
  [<b><font color="#0000FF">会员</font></b>]：当为True 时则为泡分制会员，[<b><font color="#0000FF">会员结束</font></b>]则是泡分制会员的结束时间。有关泡分制会员默认为2倍，要改修改可以修改泡分文件得到！</font></p>
<div align="center">用户资料修改程序 </div>
<form method="POST" action="showuser.asp">
  <div align="center"> 
    <select name="sjcz">
      <option value="姓名" selected>姓名</option>
      <option value="ID">ID</option>
      <option value="配偶">配偶</option>
      <option value="信箱">信箱</option>
      <option value="Oicq">OIcq</option>
    </select>
    <input type="text" name="search" size="10" maxlength="10">
    <input type="submit" value="查询" name="B1" class="p9">
    <input type="reset" name="Submit" value="清空">
  </div>
  <div align="center">ID查找一定要为数字！<br>
  </div>
  </form>
<form method="POST" action="laren.asp">
  <div align="center">拉人查看程序！<br>
    <input type="text" name="larenseek" size="10" maxlength="10">
    <input type="submit" value="查询" name="B12" class="p9">
    <input type="reset" name="Submit2" value="清空">
  </div>
  <div align="center"></div>
</form>
<form method="POST" action="pass.asp">
  <div align="center">密码修正程序<br>输入人名将他的密码改成：123456<br>
    <input type="text" name="cpass" size="10" maxlength="10">
    <input type="submit" value="修改" name="B12" class="p9">
    <input type="reset" name="Submit2" value="清空">
    <br>
  </div>
  </form>
<p align="center">『快乐江湖』</p>
</body>
</html>
