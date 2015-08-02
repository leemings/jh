<%
username=Session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
%>
<html>
<head>
<title>我的家</title>
<link rel=stylesheet href='../chatroom/css.css'>
</head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<p align=center><font color=0000ff size=5>我的爱我的梦我的家</font><br>这是心的呼唤,这是爱的交流</p><hr>
<table width=60% border=0 align=center>
<tr><td><a href='seeabout.asp' onmouseover="window.status='查看我的状态';return true;" onmouseout="window.status='';return true;" title="查看我的状态">状&nbsp;&nbsp;&nbsp;&nbsp;态</a></td></tr>
<tr><td><a href='pet.asp' onmouseover="window.status='这是心与心的呼唤,这是爱与爱的交流';return true;" onmouseout="window.status='';return true;" title="这是心与心的呼唤,这是爱与爱的交流">宠物之家</a></td></tr>
<tr><td><a href='../bbs/getpm.asp' target=rulfrm onmouseover="window.status='看看是否有朋友给自己留言';return true;" onmouseout="window.status='';return true;" title="看看是否有朋友给自己留言">我的邮箱</a></td></tr>
<tr><td><a href='seeequip.asp' onmouseover="window.status='我的武器装备';return true;" onmouseout="window.status='';return true;" title="我的武器装备">武&nbsp;&nbsp;&nbsp;&nbsp;器</a></td></tr>
<tr><td><a href='seecurative.asp' onmouseover="window.status='我的药品';return true;" onmouseout="window.status='';return true;" title="我的药品">药&nbsp;&nbsp;&nbsp;&nbsp;品</a></td></tr>
<tr><td><a href='../ballot/ballot.asp' onmouseover="window.status='参加社区投票';return true;" onmouseout="window.status='';return true;" title="参加社区投票">投&nbsp;&nbsp;&nbsp;&nbsp;票</a></td></tr>
<tr><td><a href='../other/marriage.asp' onmouseover="window.status='参加社区婚礼';return true;" onmouseout="window.status='';return true;" title="参加社区婚礼">婚&nbsp;&nbsp;&nbsp;&nbsp;礼</a></td></tr>
</table>
</body>
</html>