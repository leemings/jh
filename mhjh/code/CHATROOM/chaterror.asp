<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no cache"
Response.Expires=-1
session.Abandon
id=Trim(Request.QueryString("id"))
Select Case id
Case "000"
errormsg="本窗口即将被关闭，可能原因不明。请尝试清空INTERNET临时文件"
Case "001"
errormsg="本窗口即将被关闭。原因超时。"
Case "002"
errormsg="本窗口即将被关闭。原因是因为你被可恶的"&Request.QueryString("from")&"打死了！"
Case "003"
errormsg="本窗口即将被关闭。原因是因为你由于中毒而死亡！"
Case "004"
errormsg="你没有登录或超时断开连接，请重新登录！"
Case "005"
errormsg="本窗口即将被关闭。原因是因为你被可恶的"&Request.QueryString("from")&"驱逐出境！"
Case "006"
errormsg="本窗口即将被关闭。原因是因为你被可恶的"&Request.QueryString("from")&"逮捕入狱了，快想办法找人通融一下去吧！"
Case "007"
errormsg="本窗口即将被关闭。原因是因为可恶的"&Request.QueryString("from")&"觉得你罪大恶极，而将你投入了重刑犯专用监狱！"
Case "008"
errormsg="本窗口即将被关闭。原因是因为你的IP被锁！不会自动解除！！"
Case "009"
errormsg="本窗口即将被关闭。原因是因为你被僵尸咬死了！！"
Case else
errormsg="对不起，该出错类型未被登记。"
End Select
%>
<html>
<head><link rel="stylesheet" href="css.css" type="text/css">
<script language=javascript>if(window!=window.top)top.location.href=location.href;</script>
<title> 错 误 提 示</title>
</head>
<body oncontextmenu=self.event.returnValue=false background="bg1.gif" topmargin="150" leftmargin="200">
<table width="350" border="0" cellspacing="2" cellpadding="4">
<tr>
<td align="center">&nbsp;
<%=errormsg%></td>
</tr>
<tr>
<td align="center" valign="middle">
<input type="button"  value="确定" onClick=javascript:window.close() name="button">
</td>
</tr>
</table>
</body>
</html>
