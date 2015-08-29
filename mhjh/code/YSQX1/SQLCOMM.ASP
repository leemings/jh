<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
%>
<html>
<head>
<title>sql指令系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chatroom/css.css">
</head>
<body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif' topmargin="0">
<p align="center">快乐江湖数据库<br>
  <br>
  <b><font color="#FF0000">注意：此处操作错，将影响到数据库整个系统！<br>
  建议不明白的站长不要使用，非常危险！</font></b> </p>
<p align="left">请输入所要执行的指令：<br>
  如：update 用户 set 等级=1,身份='无' where 身份&lt;&gt;'掌门' and 门派&lt;&gt;'官府'<br>   
是将所有身份不为掌门，门派不为官府的等级清成1级,身份设置成为无！！ <br>  
  如：delete * from 物品 where 数量=0<br>   
  是将物品数据库中的物品为0的记录删除！<br>   
  如：delete * from 招式 where 门派='官府'<br>   
  是将官府的武功删除！<br>   
  如：delete * from 用户 where 等级&lt;=2 and 门派='飘花令'<br>  
是将门派为飘花令,等级小于等于2的用户删除！<br>  
  <br>  
在执行删除时，所删除记录空间并没有完全释放出来！需要下载数据库压缩！<br> 
</p> 
<div align="center">选择操作的数据库</div> 
<form method="POST" action="sqlcommok.asp"> 
  <div align="center"> 
    <br> 
    请输入查询条件：  
    <input type="text" name="sqlstr" size="50" maxlength="250"> 
    <br> 
    <input type="submit" value="执行" name="B1" class="p9"> 
    <input type="reset" name="Submit" value="清空"> 
  </div> 
  </form> 
</body> 
</html>
