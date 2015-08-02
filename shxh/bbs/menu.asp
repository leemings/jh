<%Response.Expires=-1%>
<!--#include file="set.asp"-->
<head>
<meta content='text/html; charset=gb2312' http-equiv=Content-Type>
<link rel='stylesheet' href='../chatroom/css.css'>
</head>
<body oncontextmenu='self.event.returnValue=false'>
<form name=form1>
<table width=100% bgcolor=CCCCCC><tr align=center>
<td><input type=button value=" 论坛首页 "onclick="javascript:parent.mainfrm.location.replace('main.asp');" style="border:3 bouble;background=cccccc" ></td>
<td><input type=button value=" 留 言 本 " onclick="javascript:window.open('getpm.asp','','left=100,top=50,width='+screen.availwidth*2/3+',height='+screen.availheight*3/4+',status=yes,scrollbars=yes,resizable=no');" style="border:3 bouble;background=cccccc"></td>
<td><input type=button value=" 搜寻相关 "onclick="javascript:parent.mainfrm.location.replace('searcha.asp');" style="border:3 bouble;background=cccccc"></td>
<td><input type=button value=" 退出论坛 " onclick="javascript:top.window.close();" style="border:3 bouble;background=cccccc" id=button1 name=button1></td>
</tr></table>
</form>
</body>