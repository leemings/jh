<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>江湖规矩 - happyjh.com</title>
<%
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
if sjjh_name="" then Response.Redirect "error.asp?id=440"
%>
<style type='text/css'>
body{font-size:9pt;
scrollbar-face-color:#ffffff; 
scrollbar-shadow-color:#ff9900; 
scrollbar-highlight-color:#ff9900;
scrollbar-3dlight-color:#ffffff;
scrollbar-darkshadow-color:#ffffff;
scrollbar-track-color:#ffffff;
scrollbar-arrow-color:#ff9900;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
</head>
<body>
<table border="0" width="80%" style="text-decoration: blink; color: #008080; line-height: 150%; border: 1 solid #000000" bgcolor="#FFFFFF" align="center" cellpadding="5" cellspacing="5">
  <tr> 
    <td width="100%" align="center" bgcolor="#FF9900"> <b><font color="#ffff00">1.<a href="jmhelp.asp"><font color="#ffff00">聊天室界面帮助</font></a> 
      2.</font><a href="xinshou.asp"><font color="#FFFF00">江湖入门帮助</font></a> <font color="#ffff00">3.</font><a href="jhgj.asp"><font color="#FFFF00">江湖规矩 
      </font></a><font color="#ffff00">4.</font><a href="canshu.asp"><font color="#FFFF00">江湖设定档案</font></a></b> 
    </td>
  </tr>

  <tr> 
    <td bgcolor="#FFCC00"> <p><font color="#FF0000" size="2"><b>江湖规矩</b></font> 
        <font size="2"><br>
        1.江湖中不允许骂人，说脏话。杀人不犯法，骂人要坐牢。<br>
        2.不要粘贴大量重复的话一次性发出，也不允许多次发送一句重复的话。<br>
        3.不允许在江湖中拉人</font> <font size="2">......</font> 最好不要在江湖中谈论其他的江湖(特别是站长禁制了的网址)。<br>
        上面这些基本上是会立即被系统自动踢出的。</p>
      <p>4<font size="2">.</font>其他的最好是不要威胁恐吓站长，在聊天室中别人至少有几百种方法让你倒霉<br>
        5<font size="2">.</font>不要试图做不该做的事，访问系统管理页面或进行破坏，『快乐江湖』的这些地方都设置了炸弹保护，悠着点</p>
      <p>6<font size="2">.</font>做官府的管理员一定要尽职尽责，公正无私，每个管理员的操作都会被记录在管理日志中，供大家临督。<br>
        7<font size="2">.官府人员只管理刷屏、骂人，及违反中国有关互联网规定的对象。可以使用警告、逮捕、封锁IP的手段制止恶意捣乱的行为的发生。</font> 
        <br>
        8<font size="2">.官府人员不得参与江湖恩怨，不可以偏向任何一方，只要没有捣乱行为出现，杀人的事情一般不应过问。——滥杀新手除外，欺负新朋友的人一定要加以惩罚！</font> 
        <br>
        9<font size="2">.</font><font size="2">如果官府人员很长时间不能上网，影响了自己的工作和秩序，我们会考虑选举新的管理人员来顶替。<br>
        原则上要求每天在聊天室的时间不少于3个小时。</font></p>
</td>
  </tr>
  <tr> 
    <td width="100%" align="left" style="text-indent: 0"> <p align="center"> 
      <p align="center"><a href="javascript:window.close()""><font size="4" color="#009933"><strong>我明白了</strong></font> 
        </a> </td>
  </tr>
</table>
</body>
</html>
