<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
%>
<html>
<head>
<title>用户数据库管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chatroom/css.css">
</head>
<body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif' topmargin="0">
<p align="center">魔幻江湖数据库</p><br>
<p align="center"><font color="#808000"><b>此项查询为命令查询，当命令错误、或者输入字段不存在时都会返回错误</b></font></p>
<p align="left">请输入条件查询用户：<br>
  如：身份='掌门' 查询身份为掌门的用户。<br>   
  再如： 银两&gt;10000 and 武功&gt;10000 查看银两大于10000并且武功大于10000的用户。<br>   
  在查询中可以使用：&quot;and&quot; &quot;or&quot; &quot;&gt;&quot; &quot;&lt;&quot; &quot;&lt;&gt;&quot;    
  &quot;&gt;=&quot; &quot;&lt;=&quot; &quot;=&quot; 关系量！<br>   
在查询物品时：如输入：所有者='小宇'<br> 
  在查询武功数据库时：如输入：姓名='小宇'<br> 
   <font color="#0000FF">查询字段：此值为单一值,如：内力、体力等。以下为错误：内力 and 体力 或 内力,体力等,</font>默认值为id</p>  
<div align="center">用户资料修改程序 </div>  
<form method="POST" action="seeuserok.asp">  
  <div align="center"><font color="#FFFFFF">  
    <select name="seekfs">  
      <option value="0" selected>查询用户数据库</option>  
      <option value="1">查询物品数据库</option>  
      <option value="2">查询武功数据库</option>  
    </select>  
    </font><br>  
    <br>  
    请输入查询条件：   
    <input type="text" name="tiaojian" size="50" maxlength="250">  
    <br>  
    输入将查询字段：   
    <input type="text" name="show" value="id" size="10" maxlength="12">  
    <br>  
    <input type="submit" value="查询" name="B1" class="p9">  
    <input type="reset" name="Submit" value="清空">  
  </div>  
  </form>  
</body>  
</html>