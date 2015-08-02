<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
%>
<html>
<head>
<title><%=Application("Ba_jxqy_systemname")%></title>
<link rel=stylesheet href=chatroom/css.css>
</head>
<body  oncontextmenu=self.event.returnValue=false bgcolor='<%=bgcolor%>' background='<%=bgimage%>'>
<table  bgcolor="#FFE4CA" width=75% align=center border=1 bordercolorlight=000000 cellspacing=0 cellpadding=1 bordercolordark=FFFFFF >
<tr align=center><td><a href='#' onclick="javascript:mainwin=window.open('main.asp','main','left=0,top=0,status=yes,scrollbars=yes,resizable=yes');mainwin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='使用地图导航系统';return true;" onmouseout="window.status='';return true;" title="使用地图导航系统">地图导航</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='myhome/myhome.asp' target=rulfrm onmouseover="window.status='我的爱我的梦我的家';return true;" onmouseout="window.status='';return true;" title="我的爱我的梦我的家">我 的 家</a></td></tr>
<tr align=center><td><a href='chatroom/announce.htm' target=rulfrm onmouseover="window.status='观看官府临时公告';return true;" onmouseout="window.status='';return true;" title="观看官府临时公告">会员须知</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF" ><a href='#' onclick="javascript:admwin=window.open('admin/administrator.asp','admin','left=0,top=0,status=yes,scrollbars=yes,resizable=yes');admwin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='进入官府';return true;" onmouseout="window.status='';return true;" title="进入官府">官&nbsp;&nbsp;&nbsp;&nbsp;府</a></td></tr>
<tr align=center><td><a href='#' onclick="javascript:chatwin=window.open('chatroom/chatroom.asp','_blank');" onmouseover="window.status='进入聊天室';return true;" onmouseout="window.status='';return true;" title="进入聊天室">聊 神 谷</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='#' onclick="javascript:bbswin=window.open('bbs/index.asp','bbs','left=0,top=0,status=yes,scrollbars=yes,resizable=yes');bbswin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='进入聚贤山庄';return true;" onmouseout="window.status='';return true;" title="进入聚贤山庄">聚贤山庄</a></td></tr>
<tr align=center><td ><a href='#' onclick="javascript:advwin=window.open('adventure/adventure.asp','adventure','left=100,top=50,status=yes,scrollbars=yes,resizable=yes');advwin.resizeTo(screen.availWidth*2/3,screen.availHeight*3/4);" onmouseover="window.status='开始你的冒险生涯';return true;" onmouseout="window.status='';return true;" title="开始你的冒险生涯">探 险 谷</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='street/street.asp' target=rulfrm onmouseover="window.status='商旅云集行人接踵处';return true;" onmouseout="window.status='';return true;" title="在这儿买卖物品">十里长街</a></td></tr>
<tr align=center><td ><a href='other/dubo.asp' target=rulfrm onmouseover="window.status='十赌九诈,小心为妙';return true;" onmouseout="window.status='';return true;" title="十赌九诈,小心为妙">赌&nbsp;&nbsp;&nbsp;&nbsp;场</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='#' onclick="javascript:stowin=window.open('stock/stock.htm','stock','left=0,top=0,status=yes,scrollbars=yes,resizable=yes');stowin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='从事证券买卖的场所';return true;" onmouseout="window.status='';return true;" title="从事证券买卖的场所">证 交 所</a></td></tr>
<tr align=center><td><a href='street/hotel.asp' target=rulfrm onmouseover="window.status='累了,我要睡觉了';return true;" onmouseout="window.status='';return true;" title="累了,我要睡觉了">悦来客栈</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='other/top.asp' target=rulfrm onmouseover="window.status='传为百晓生作,祸乱源泉';return true;" onmouseout="window.status='';return true;" title="传为百晓生作,祸乱源泉">兵 器 谱</a></td></tr>
<tr align=center><td><a href='street/herostele.asp' target=rulfrm onmouseover="window.status='让光荣与梦想随鹰背苍茫远去';return true;" onmouseout="window.status='';return true;" title="让光荣与梦想随鹰背苍茫远去">英 烈 堂</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='suicide.asp' target=rulfrm onmouseover="window.status='我的梦想在天的那一边';return true;" onmouseout="window.status='';return true;" title="我的梦想在天的那一边">断 魂 涯</a></td></tr>
<tr align=center><td><a href='#' onclick="javascript:chatwin=window.open('diaoyu/diaoyu.asp','chatroom','left=0,top=0,status=yes,scrollbars=yes,resizable=yes');chatwin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='进入聊天室';return true;" onmouseout="window.status='';return true;" title="进入聊天室">钓 
    鱼 台</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='#' onclick="javascript:chatwin=window.open('Qmg/Qmg.htm','chatroom','left=0,top=0,status=yes,scrollbars=yes,resizable=yes');chatwin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='进入聊天室';return true;" onmouseout="window.status='';return true;" title="进入聊天室">巧 
    面 馆</a></td></tr>
<tr align=center><td bgcolor="#FFE4CA"><a href='#' onclick="javascript:chatwin=window.open('jiudian/jd.asp','chatroom','left=0,top=0,status=yes,scrollbars=yes,resizable=yes');chatwin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='进入聊天室';return true;" onmouseout="window.status='';return true;" title="进入聊天室">大 
    酒 店</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='#' onclick="javascript:chatwin=window.open('hktk/xuetang.asp','chatroom','left=0,top=0,status=yes,scrollbars=yes,resizable=yes');chatwin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='进入聊天室';return true;" onmouseout="window.status='';return true;" title="进入聊天室">训 
    练 室</a></td></tr>
<tr align=center><td bgcolor="#FFE4CA"><a href='#' onclick="javascript:chatwin=window.open('myjinmao/index.asp','chatroom','left=0,top=0,status=yes,scrollbars=yes,resizable=yes');chatwin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='进入聊天室';return true;" onmouseout="window.status='';return true;" title="进入聊天室">抗击伪寇</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='#' onclick="javascript:chatwin=window.open('yhy/index.asp','chatroom','left=0,top=0,status=yes,scrollbars=yes,resizable=yes');chatwin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='进入聊天室';return true;" onmouseout="window.status='';return true;" title="进入聊天室">江湖妓院</a></td></tr>
<tr align=center><td bgcolor="#FFE4CA"><a href='welcome.asp' target=rulfrm onmouseover="window.status='重返首页';return true;" onmouseout="window.status='';return true;" title="重返首页">返&nbsp;&nbsp;&nbsp;&nbsp;回</a></td></tr>
<tr align=center><td bgcolor="#FFFFFF"><a href='exit.asp' target=rulfrm onmouseover="window.status='退出并重新进入';return true;" onmouseout="window.status='';return true;" title="退出并重新进入">退&nbsp;&nbsp;&nbsp;&nbsp;出</a></td></tr>
</table>
</body>
</html>