<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<html>
<head>
<title>江湖赌场</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style type="text/css"><!--td           { font-family: 宋体; font-size: 9pt }
body         { font-family: 宋体; font-size: 9pt }
select       { font-family: 宋体; font-size: 9pt }
a            { text-decoration: none; font-family: 宋体; font-size: 9pt }
a:hover      { text-decoration: underline; color: #CC0000; font-family: 宋体; font-size:
9pt }
.big         { font-family: 宋体; font-size: 12pt }
--></style>
</head>

<body text="#000000" background="../jhimg/bk_hc3w.gif">
<br>
<table width="633" align="center" cellspacing="2" border="2" cellpadding="5"
bgcolor="#90c088">
<tr>
<td style="font-size: 21; color: #FF1493" align="center" height="27"
bgcolor="#336633" width="613"><b>江 湖 赌 街</b></td>
</tr>
<tr>
<td width="613" height="171">
<table border="0" cellspacing="1" cellpadding="2" width="100%"
bgcolor="#000000" bordercolorlight="#EFEFEF" height="95">
<tr bgcolor="#DFEDFD">
<td width="12%" height="47">
<div align="center"> <a href="jzstb\index.asp">剪刀石头布</a> </div>
</td>
<td width="18%" height="47">
<div align="center"> <a href="dice.asp">赌 骰 子</a> </div>
</td>
<td width="17%" height="47">
<div align="center"> <a href="b&amp;s.asp">赌 大 小</a> </div>
</td>
          <td width="11%" align="center" height="47"><a
class=blue href="#" onClick="window.open('sama/horse.asp','aqjhhorse','Status=yes,scrollbars=yes,resizable=yes,fullscreen=yes')" ><font size="2">江湖马场</a></td>
<td width="11%" align="center" height="47"><a href="21/index.asp"
target="_blank">21点</a></td>
</tr>
<tr bgcolor="#f7f7f7">
<td width="12%" height="48">猜中每次赢50两银子，输一次失银子50两</td>
<td width="18%" height="48">庄家随机出三个<font class="c"><b>骰子</b>，大侠您也随机出三个<b>骰子</b>，谁大，谁赢，赔率１：１</font>
<td width="17%" height="48">庄家随机出三个<font class="c"><b>骰子</b>，你选择开大还是小，如果<b>骰子</b>加起来大于等于１０为大，反之为小，赔率１：１</font>
<td width="11%" height="48"><font class="c">动态跑马您自己看咯：）</font>
          <td width="11%" height="48">抓大点要看运看，更要看手气！
        </tr>
</table>
<br>
堵场混混：欢迎大侠来玩，如果哪位不知趣抽老千，让我们老大发现了，一事实上打你小子满地找牙！</td>
</tr>
</table>
<div align="center"><br>版权所有≮快乐江湖总站≯</div>
</body>
</html>
