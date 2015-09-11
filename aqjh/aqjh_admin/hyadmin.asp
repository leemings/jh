<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
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
<p align="center"><b><font color="#FF0000">注意：</font></b><font color="#FF0000">对于已经是会员的，设置完成后时间会按会员结束时间加<br>
  上设置时间，会员等级会设置成新的等级!</font></p>
<p align="center"><font color="#000000">说明：等级制会员，见下面说明！泡点成倍制会员，为图形版新增功能！设置成会员后<br>
  他在聊天中泡点为正常的N倍 ，N可以在paodian.asp中设置！</font><font color="#FF0000"><br>
  <br>
  (<font color="#0000FF"><b>泡点成倍制会员不分等级！</b></font>)<br>
  </font><a href="hydj.asp"><b>等级制会员</b></a></p>
<p align="center"><a href="hypd.asp"><b>泡点制会员</b></font></a></p>
<p align="center"><a href="hylist.asp">现有会员列表</a> <a href="fine.asp">更新用户资料</a> <a href="kfaddadmin.asp">会员批量送卡 </a> <a href="jiangli.asp"><font color="#FF0000">发放奖励</font></a></p>
</body></html>