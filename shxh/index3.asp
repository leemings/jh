<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>页脚</title>

<link rel="stylesheet" href="chatroom/css1.css" type="text/css">
<body background="images/bg.gif" text="#303030" topmargin="1" leftmargin="0" bgcolor="#FFE1C4">
<table border="0" width="100%" align="center" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
  <tr align="center"> 
    <td>论剑亭</td>
    <td><a href="#" onclick="javascript:window.open('adventure/adventure.asp','adventure','toolbar=0,location=0,status,menubar=0,left=100,top=50,width='+screen.availWidth*2/3+',Height='+screen.availHeight*3/4)">探险谷</a></td>
    <td><a href="street/market.asp" target="main">集市</a></td>
    <td><a href="street/armorshop.asp" target="main">武具</a></td>
    <td><a href="street/curativeshop.asp" target="main">药店</a></td>
    <td><a href="street/pawnshop.asp" target="main">当铺</a></td>
    <td><a href="other/dubo.asp" target="main">赌场</a></td>
    <td><a href=# onclick="javascript:window.open('stock/stock.htm','stock','left=0,top=0,status,resizable=0,width='+screen.availWidth+',height='+screen.availHeight)">证交所</a></td>
    <td><a href="street/cardshop.asp" target="main">道具店</a></td>
    <td><a href="index2.asp" target="main">虚幻地图</a></td>
    <td><a href="exit.asp" target="_top">退出虚幻</a></td>
  </tr>
</table>
<div align="center"><br>
  程序版权：<a href="http://www.mxcz.org/" target="_blank"><font color="#990000"> 
  <% response.write Application("Ba_jxqy_copyrightassert")%>
  </font></a> <font color="#990000"></font> 
  &nbsp;&nbsp; 授权给： <font color="#990000"> 
  <% response.write Application("Ba_jxqy_userright")%>
  &nbsp;&nbsp; </font>软件序列号： <font color="#990000"> 
  <% response.write Application("Ba_jxqy_seriesnumber")%>
  </font>
  访问统计：<font color="#993300">  
          <%=Application("Ba_jxqy_visitor")%>  
          </font> 会员功能：<font color="#993300">  
          <%if Application("Ba_jxqy_fellow")=true then%>
          开放
          <%else%>
          关闭
          <%end if%>
          </font>  
</div>
</body>
</html>

