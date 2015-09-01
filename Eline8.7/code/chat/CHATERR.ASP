<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
id=Trim(Request.QueryString("id"))
if id="001" or id="002" then
	if sjjh_name<>"" then
	to1=sjjh_name
	mycd=DateDiff("n",Session("sjjh_savetime"),now())
	if mycd>50 then mycd=50
	if mycd<1 then mycd=1
	Application.Lock
	Application("sjjh_diuqi")="v|"&to1&"|"&mycd&"|"&now()
	Application.UnLock
	diuqi=sjjh_name&"掉出聊天室，尚未保存存点总计:<a href='dq.asp' target='d'><b><font color=red>+"&mycd&"</font></b></a>点,谁抢到是谁的…"
	towhoway=0
	saycolor="660099"
	addwordcolor="660099"
	addsays="对"
	says1="<b>【江湖消息】</b>"&diuqi
	says1=replace(says1,"'","\'")
	says1=replace(says1,chr(34),"\"&chr(34))
	saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & "踢出" & chr(39) &","& chr(39) & to1 & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & to1 & chr(39) &"," & chr(39) & says1 & chr(39) &",0," & session("nowinroom") & ");</script>"
	addmsg saystr
	Function Yushu(a)
		Yushu=(a and 31)
	End Function
	Sub AddMsg(Str)
		Application.Lock()
		Application("SayCount")=Application("SayCount")+1
		i="SayStr"&YuShu(Application("SayCount"))
		Application(i)=Str
		Application.UnLock()
	End Sub
	end if
end if
%>
<!--
if id="003" then
	if sjjh_name<>"" then
	to1=sjjh_name
	mycd=DateDiff("n",Session("sjjh_savetime"),now())
	if mycd>50 then mycd=50
	if mycd<1 then mycd=1
	Application.Lock
	Application("sjjh_diuqi")="v|"&to1&"|"&mycd&"|"&now()
	Application.UnLock
	diuqi=sjjh_name&"<font color=#000000>试图一机多号，被系统请出了聊天室！大家快抢存点吧:</font><a href='dq.asp' target='d'><b><font color=red> +"&mycd&" </font></b></a><font color=#000000>点</font>"
        act="消息"
	towhoway=0
	saycolor="660099"
	addwordcolor="660099"
	addsays="对"
	says1="<font color=#000000><b>〖江湖密报〗</b></font>"&diuqi
	says1=replace(says1,"'","\'")
	says1=replace(says1,chr(34),"\"&chr(34))
	saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & "act" & chr(39) &","& chr(39) & to1 & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & to1 & chr(39) &"," & chr(39) & says1 & chr(39) &",0," & session("nowinroom") & ");</script>"
	addmsg saystr
	Function Yushu(a)
		Yushu=(a and 31)
	End Function
	Sub AddMsg(Str)
		Application.Lock()
		Application("SayCount")=Application("SayCount")+1
		i="SayStr"&YuShu(Application("SayCount"))
		Application(i)=Str
		Application.UnLock()
	End Sub
	end if
end if
-->
<%
Session.Abandon
Select Case id
Case "000"
	nl="　　对不起，程序所在目录不是虚拟目录，或设置有错误，Global.asa 没有被执行。本程序需要虚拟目录的支持！"
Case "001"
	nl="本窗口即将被关闭。<br><br>造成此现象的原因主要有：<br>　　１由于网络传输问题，导致你的机器与服务器在三分钟内无法传递数据，系统将你作为超时而断开了连接；<br>　　２你点击了“换名登录”重回登录页面，又没有先关闭此窗口；<br>　　３你被踢出了聊天室。<br>　　４如果你使用本窗口进入不会出现这个问题，而用新窗口进入就出现这个问题的话，问题就出在你的机器：对弹出的新窗口无法继承上级窗口的值。<br><br>解决方法：<br>　　关闭此窗口，重新到登录页面输入用户名和密码进行登录。如果你是被踢出聊天室，则你方才所用的用户名在 5 分钟内不能使用。如果是第４种情况，可以试着清除ＩＥ的临时文件，然后重新启动机器。"
Case "002"
	nl="本窗口即将被关闭。<br><br>原因：<br>　　你超过 " & Application("sjjh_maxtimeout") & " 分钟没有发言，为减轻服务器负担，系统自动清除你所占用的资源。<br><br>解决方法：<br>　　关闭此窗口，重新到登录页面输入用户名和密码进行登录。"
Case "003"
	nl="本窗口即将被关闭<br><font color=blue><b>请不要一台计算机上多个帐号！</b></font><br><br>造成此现象的原因主要有:<br>为了节省服务器资源，我们规定一台计算机只可以上一个帐号，当你上了一个号时，如果再上第二个号那么第一个号会被强行关闭。<br><br>解决办法:<br>一台计算机只可以上一个帐号<br><font color=blue><b>刷新登陆页面</font>重新登陆！"
Case else
	nl="　　对不起，该出错类型未被登记。"
End Select
nl="<p>" & nl & "</p>"%><html>
<head>
<title>出错提示</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor=<%=chatroombgcolor%> background=<%=chatroombgimage%>>
<script LANGUAGE="JavaScript">if(window!=window.top)top.location.href=location.href;</script>
<table width="100%" border="0" height="100%">
<tr align="center">
<td>
<form method="post" action="">
<table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" bgcolor="E0E0E0">
<tr>
<td>
<table border="0" bgcolor="#FF0099" cellspacing="0" cellpadding="2" width="350">
<tr>
<td width="342"><font color="FFFFFF">□出错提示</font></td>
<td width="18">
<table border="1" bordercolorlight="666666" bordercolordark="FFFFFF" cellpadding="0" bgcolor="E0E0E0" cellspacing="0" width="18">
<tr>
<td width="16"><b><a href="javascript:window.close()" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="关闭"><font color="000000">×</font></a></b></td>
</tr>
</table>
</td>
</tr>
</table>
<table border="0" width="350" cellpadding="4">
<tr>
<td width="59" align="center" valign="top"><font face="Wingdings" color="#FF0000" style="font-size:32pt">L</font></td>
<td width="269">
<%=nl%>
</td>
</tr>
<tr>
<td colspan="2" align="center" valign="top">
<input type="button" name="ok" value="　确 定　" onclick=javascript:window.close()>
</td>
</tr>
</table>
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
<script Language="JavaScript">
document.forms[0].ok.focus();
</script>
</body>
</html>