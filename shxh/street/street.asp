<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
set conn=server.CreateObject("adodb.connection")
conn.open Application("BA_jxqy_connstr")
conn.execute "update 系统设置 set 属性值='"&Application("Ba_jxqy_visitor")&"' where 属性='访问人数'"
conn.close
set conn=nothing
%>
<html>
<head>
<title><%=Application("Ba_jxqy_systemname")%></title>
<link rel=stylesheet href=../chatroom/css.css>
</head>
<body  oncontextmenu=self.event.returnValue=false bgcolor='<%=bgcolor%>' background='<%=bgimage%>'>
<p align=center><font color=0000ff size=5>十里长街</font></p><hr>
<table  bgcolor="#FFE4CA" width=30% align=center border=1 bordercolorlight=000000 cellspacing=0 cellpadding=1 bordercolordark=FFFFFF >
<tr align=center><td ><a href='../other/factory.asp' onmouseover="window.status='流自己的汗,吃自己的饭,靠天靠地靠祖宗,不算是好汉';return true;" onmouseout="window.status='';return true;" title="流自己的汗,吃自己的饭,靠天靠地靠祖宗,不算是好汉">打 工 坊</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='bank.asp' target=rulfrm onmouseover="window.status='在这儿存钱取钱';return true;" onmouseout="window.status='';return true;" title="在这儿存钱取钱">钱&nbsp;&nbsp;&nbsp;&nbsp;庄</a></td></tr>
<tr align=center><td ><a href='armorshop.asp' target=rulfrm onmouseover="window.status='在这儿购买头盔,盔甲,武器,防具,饰品,暗器';return true;" onmouseout="window.status='';return true;" title="在这儿购买头盔,盔甲,武器,防具,饰品,暗器">武 器 店</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='curativeshop.asp' target=rulfrm onmouseover="window.status='在这儿购买药品';return true;" onmouseout="window.status='';return true;" title="在这儿购买药品">山 药 店</a></td></tr>
<tr align=center><td><a href='market.asp' target=rulfrm onmouseover="window.status='在这儿和大家自由买卖自己不用的物品';return true;" onmouseout="window.status='';return true;" title="在这儿和大家自由买卖自己不用的物品">自由集市</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='pawnshop.asp' target=rulfrm onmouseover="window.status='要来吗?不用担心,秦琼也有卖马时!';return true;" onmouseout="window.status='';return true;" title="要来吗?不用担心,秦琼也有卖马时!">当&nbsp;&nbsp;&nbsp;&nbsp;铺</a></td></tr>
<tr align=center><td ><a href='hotel.asp' target=rulfrm onmouseover="window.status='累了,我要睡觉了';return true;" onmouseout="window.status='';return true;" title="累了,我要睡觉了">悦来客栈</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='cardshop.asp' target=rulfrm onmouseover="window.status='这儿有道具可买';return true;" onmouseout="window.status='';return true;" title="道具店">道 具 店</a></td></tr>
</table>
</body>
</html>