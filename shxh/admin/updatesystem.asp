<%
Response.Buffer=true
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
if not(session("Ba_jxqy_usercorp")="官府" and Session("Ba_jxqy_usergrade")=Application("Ba_jxqy_allright")) then Response.Redirect "../error.asp?id=500"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open application("Ba_jxqy_connstr")
Set rst=Server.CreateObject("ADODB.RecordSet")
rst.Open "系统设置",conn,1,2
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
i=1
do while not rst.EOF 
	msg=msg&"<tr><td width='30%'>"&rst("属性")&"</td><td>"&Request.Form(i)&"</td></tr>"
	rst("属性值")=Request.Form(i)
	rst.Update
	rst.MoveNext
	i=i+1
loop
rst.Requery
do while Not rst.EOF
		Select Case rst("属性")
			Case "系统名称"
				Application("Ba_jxqy_systemname")=rst("属性值")
			Case "用户名"
				Application("Ba_jxqy_userright")=rst("属性值")
			Case "版权声明"
				Application("Ba_jxqy_copyrightassert")=rst("属性值")
			Case "序列号"
				Application("Ba_jxqy_seriesnumber")=rst("属性值")
			Case "背景色"
				Application("Ba_jxqy_backgroundcolor")=rst("属性值")	
			Case "背景图片"
				Application("Ba_jxqy_backgroundimage")=rst("属性值")	
			Case "序列号"
				Application("Ba_jxqy_backgroundcolor")=rst("属性值")	
			Case "访问人数"
				Application("Ba_jxqy_visitor")=rst("属性值")
			Case "最近开放时间"
				rst("属性值")=cstr(now())
				rst.Update
				Application("Ba_jxqy_opendata")=cstr(now())
			Case "积分设置"
				Ba_grade=split(rst("属性值"),";")
				for i=0 to ubound(Ba_grade)
					application("Ba_jxqy_grade"&i)=clng(Ba_grade(i))
				next
			Case "进入提示"
				Application("Ba_jxqy_guestjoin")=rst("属性值")
			Case "退出提示"
				Application("Ba_jxqy_guestleft")=rst("属性值")
			Case "断线提示"
				Application("Ba_jxqy_guestisnotconnection")=rst("属性值")
			Case "踢人权利"
				Application("Ba_jxqy_kickguestright")=clng(rst("属性值"))	
			Case "逮捕权利"
				Application("Ba_jxqy_arrestright")=clng(rst("属性值"))
			Case "入狱权利"
				Application("Ba_jxqy_gaolright")=clng(rst("属性值"))
			Case "封锁权利"
				application("Ba_jxqy_lockipright")=clng(rst("属性值"))		
			Case "炸人权利"
				Application("Ba_jxqy_bombright")=clng(rst("属性值"))
			Case "解锁权利"
				Application("Ba_jxqy_unlockipright")=clng(rst("属性值"))
			Case "升级权利"
				Application("Ba_jxqy_exaltgraderight")=clng(rst("属性值"))
			Case "降级权利"
				Application("Ba_jxqy_declinegraderight")=clng(rst("属性值"))
			Case "站长权利"
				Application("Ba_jxqy_allright")=clng(rst("属性值"))				
			Case "禁用名称"
				Application("Ba_jxqy_illegidimacyname")=rst("属性值")
			Case"允许不发言时间"
				Application("Ba_jxqy_nosaytime")=clng(rst("属性值"))
			Case "系统密钥"
				Application("Ba_jxqy_passwordkey")=rst("属性值")
			case "公告"	
				Application("Ba_jxqy_advertisemen")=rst("属性值")
			case "公告高度"
				Application("Ba_jxqy_advertisemenheight")=clng(rst("属性值"))
			case "泡点设置"
				Application("Ba_jxqy_paycent")=clng(rst("属性值"))
			case "允许在线"
				Application("Ba_jxqy_maxonline")=clng(rst("属性值"))
			Case "新人积分"
				Application("Ba_jxqy_newuser")=clng(rst("属性值"))
			case "会员功能开放"
				Application("Ba_jxqy_fellow")=cbool(rst("属性值"))
			Case "最大攻击"
				Application("Ba_jxqy_Maxattack")=clng(rst("属性值"))
			Case "赌场间隔时间(秒)"
				Application("Ba_jxqy_bettime")=clng(rst("属性值"))			
		end select
		rst.MoveNext
	loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head>
<link rel="stylesheet" href="../chatroom/css.css">
</head>
<body oncontextmenu=self.event.returnValue=false topmargin=0 background="<%=bgimage%>" bgcolor="<%=bgcolor%>">
<table width=100% border=3><tr><td>系统属性</td><td>属性新值</td></tr>
<%=msg%>
</table>
<p align=center>系统更新成功<br>
<input type="button" name="ok" value="　确 定　" onclick=javascript:history.go(-1)></p>
</body></html>