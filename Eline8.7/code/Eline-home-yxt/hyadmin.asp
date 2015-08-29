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
<title>会员数据库管理♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>
<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif">
<p align="center"><%=Application("sjjh_chatroomname")%>数据库</p>
<p align="center"><a href="hygl.asp">会员在线申请管理</a></p>
<p align="center"><b><font color="#FF0000">注意：</font></b><font color="#FF0000">对于已经是会员的，设置完成后时间会按会员结束时间加<br>
  上设置时间，会员等级会设置成新的等级!</font></p>
<p align="center"><font color="#000000">说明：等级制会员，见下面说明！泡点成倍制会员，为6.2版新增功能！设置成会员后<br>
  他在聊天中泡点为正常的N倍 ，N可以在paodian.asp中设置！</font><font color="#FF0000"><br>
  <br>
  (<font color="#0000FF"><b>泡点成倍制会员不分等级！</b></font>)<br>
  </font><a href="hydj.asp"><font size="+2" face="楷体_GB2312"><b>等级制会员</b></font></a></p>
<p align="center"><a href="hypd.asp"><font size="+2" face="楷体_GB2312"><b>泡点制会员</b></font></a></p>
<p align="center"><a href="../jhmp/hy.asp">现有会员列表</a> <a href="fine.asp">更新用户资料</a></p>
</body>
</html>
