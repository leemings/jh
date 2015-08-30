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
<title>控制中心♀wWw.happyjh.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>

<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif">
<p><font size="5" face="楷体_GB2312" color="#3333FF">首先感谢您使用</font><font color="#0000FF" size="5" face="楷体_GB2312"><b><i>『快乐江湖』</i></b></font></p>
<p>本江湖原始版本借用<font color="#0000FF">世纪江湖</font>开发而成，<br>
  在此，特感谢所有支持我们江湖发展的朋友，及提出意见的网友！</p>
<p>本江湖采用多种加密技术，因为不要修改本程序如：盗版、随意修改程序等</p>
<p>为此，对于本程充产生的任何后果将不负任何法律上的责任！</p>
<p>最后谢谢您的使用！</p>
<p>本套程序采用了数据库加密，唯一网管理等新技术防黑，是现在最安全、速度最快的江湖！</p>
<p align="center"><font color="#000000" size="-1">【 版权：</font><font color="#0000FF" face="楷体_GB2312"><b><i>快乐江湖</i></b></font><font color="#000000" size="-1"> 授权：<%=Application("sjjh_user")%>  
  运行序列号：<%=Application("sjjh_sn")%> 站长：<%=Application("sjjh_admin")%> 】</font></p> 
</body> 
</html> 
