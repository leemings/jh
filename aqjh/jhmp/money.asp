<html>
<head>
<%
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');window.close();</script>"
	Response.End
end if
%>
<title>工资领取处</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>
TD {FONT-FAMILY: "宋体"; FONT-SIZE: 9pt}
BODY {FONT-FAMILY: "宋体"; FONT-SIZE: 9pt}
SELECT {FONT-FAMILY: "宋体"; FONT-SIZE: 9pt}
A {COLOR: #FFC106; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: none}
A:hover {COLOR: #cc0033; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
</STYLE>
</head>
<body bgcolor="#CCCCCC" text="#000000" leftmargin="0" background="../bg.gif">
<div align="center">
<p><b><img border="0" src="../chat/img/menoy.gif"><font color="#008080"><b>工资领取处</b></font></b> <br>
<font color="#800000"><font size="2">会员每天可以领到一定的银两和会员金卡,等级越高所得金卡越多；掌门每天可领到一定的金币，玩家战斗等级20级以上可领取金币1粒。会员掌门请到会员领钱处领取工资,薪水计算方法首次带入离派及转生次数计算，每天一次都能领到1点工资小点，用处非常大。</font>
</font>
</p>
<p><a href="MONEY1.ASP"><font color="#008000">一级等级制会员</font></a>
</p>
<p><a href="MONEY2.ASP"><font color="#008000">二级等级制会员</font></a>
</p>
<p><a href="MONEY3.ASP"><font color="#008000">三级等级制会员</font></a>
</p>
<p><a href="MONEY4.ASP"><font color="#008000">四级等级制会员</font></a>
</p>
<p><a href="MONEY10.ASP"><font color="#008000">五级等级制会员</font></a>
</p>
<p><a href="MONEY06.ASP"><font color="#008000">六级等级制会员</font></a>
</p>
<p><a href="MONEY07.ASP"><font color="#008000">七级等级制会员</font></a>
</p>
<p><a href="MONEY08.ASP"><font color="#008000">八级等级制会员</font></a>
</p>
<p><a href="MONEY5.ASP"><font color="#008000">泡点制会员领钱</font></a>
</p>
<p><a href="MONEY6.ASP"><font color="#008000">掌门区</font></a><font color="#008000">-</font><a href="MONEY7.ASP"><font color="#008000">官府区</font></a>
</p>
<p><a href="MONEY8.ASP"><font color="#008000">普通人领钱</font></a>
</p>
<p><a href="MONEY9.ASP"><font color="#008000">20以上非掌门</font></a>
</p>
<p align="center">
<input type=button value=关闭窗口 onClick='window.close()' name="button">
</p>
</div>
</body>
</html>