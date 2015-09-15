<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
if aqjh_name<>Application("aqjh_user") then Response.Redirect "../error.asp?id=486"
%>
<!--#include file="../config.asp"-->
<html>
<head>
<title>系统参数</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=css/css.css type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<%if not request.form("set")="ok" then%>
<div align="center"> 【系统参数】</div>
<hr noshade size="1" color=009900>
<b>［提示］</b>注意，所有参数必填不可为空，不可以用全角字符输入参数。
<a href="javascript:history.go(0)">刷新</a> 
<hr noshade size="1" color=009900>
<form name="form1" method="post" action="">
<input type="hidden" name="aqjh_chatbgcolor" value="<%=Session("afa_chatbgcolor")%>">
<input type="hidden" name="aqjh_chatimage" value="<%=Application("aqjh_chatimage")%>">
<input type="hidden" name="aqjh_chatcolor" value="<%=Application("aqjh_chatcolor")%>">
<input type="hidden" name="aqjh_iplocktime" value="<%=Application("aqjh_iplocktime")%>">
<p><font color=red>************ 江湖其它参数***********</font><br>
<input type="text" name="aqjh_maxtimeout" value="<%=Application("aqjh_maxtimeout")%>">
泡点超时时间，不要设置太长，否则拖累服务器[xx分钟]<br>
<input type="text" name="aqjh_maxnpc" value="<%=Application("aqjh_maxnpc")%>">
npc总在线人数<br>
<input type="text" name="aqjh_npcwp" value="<%=application("aqjh_npcwp")%>">
npc暴涨比较值[建议不要改动]<br>
<input type="text" name="aqjh_npcff" value="<%=application("aqjh_npcff")%>">
npc发放需要的等级[管理级,建议用默认10级管理]<br>
<input type="text" name="aqjh_baowuyin" value="<%=Application("aqjh_baowuyin")%>">
每一级修练宝物所得到的银两<br>
<input type="text" name="aqjh_baowuxl" value="<%=Application("aqjh_baowuxl")%>">
每次修练的级数，如100级的需修练4次！他所得到的上限就是4！这个值是用战斗等级取整的！<br>
<input type="text" name="aqjh_killman" value="<%=Application("aqjh_killman")%>">
江湖一天的杀人数，夺宝一次算一次杀人<br>
<input type="text" name="aqjh_wgsx" value="<%=aqjh_wgsx%>">
每升一级的武功加值,此值当aqjh_sx=1时才有作用<br>
<input type="text" name="aqjh_nlsx" value="<%=aqjh_nlsx%>">
每升一级的内力加值,此值当aqjh_sx=1时才有作用<br>
<input type="text" name="aqjh_tlsx" value="<%=aqjh_tlsx%>">
每升一级的体力加值,此值当aqjh_sx=1时才有作用<br>
<input type="text" name="aqjh_gjsx" value="<%=aqjh_gjsx%>">
每升一级的攻击加值,此值当aqjh_sx=1时才有作用<br>
<input type="text" name="aqjh_fysx" value="<%=aqjh_fysx%>">
每升一级的防御加值,此值当aqjh_sx=1时才有作用<br>
<input type="hidden" name="aqjh_chat_maxpeople" value="<%=Application("aqjh_chat_maxpeople")%>" readonly>
<p><font color=red>************ 江湖开关设置***********</font><br>
<input type="text" name="aqjh_closedoor" value="<%=Application("aqjh_closedoor")%>">
聊天室关门,0为开门,1为关门<br>
<input type="text" name="aqjh_baowu" value="<%=Application("aqjh_baowu")%>">
是否有宝物当为0时没有宝物！宝物出现是会影响聊天室速度的，服务器不好关闭！<br>
<input type="text" name="aqjh_sx" value="<%=aqjh_sx%>">
是否有上限0没有上限，1有上限<br>
<input type="text" name="aqjh_automanname" value="<%=Application("aqjh_automanname")%>">
聊神名称设置<br>
<input type="text" name="aqjh_userout" value="<%=Application("aqjh_userout")%>">
退出聊天室的信息(注意里面的"一定要为全角字符)<br>
<input type="text" name="aqjh_dieip" value="<%=aqjh_dieip%>">
ip永久封锁使用前按格式书写，多个用;号隔开,后面请保留有分号<br>
<input type="text" name="aqjh_baowuname" value="<%=Application("aqjh_baowuname")%>">
宝物名及宝物说明<br>
<input type="text" name="aqjh_baowusm" value="<%=Application("aqjh_baowusm")%>">
得到宝物后的消息 <br>
<hr noshade size="1" color=009900>
<p><font color=red>************ 江湖系统设置***********</font><br>
<input type="hidden" name="aqjh_chatroomname" value="<%=Application("aqjh_chatroomname")%>" readonly>
<input type="text" name="aqjh_homepage" value="<%=Application("aqjh_homepage")%>">
江湖域名<br>
<input type="hidden" name="aqjh_sn" value="<%=Application("aqjh_sn")%>" readonly>
<input type="text" name="aqjh_qq" value="<%=Application("aqjh_qq")%>">
站长qq<br>
<input type="text" name="aqjh_email" value="<%=Application("aqjh_email")%>">
站长email<br>
<input type="hidden" name="aqjh_user" value="<%=Application("aqjh_user")%>" readonly>
<input type="text" name="aqjh_admin" value="<%=Application("aqjh_admin")%>">
允许进入的10管理员:多个之间以半角逗号,分隔如：张三,李四<br>
<input type="text" name="hidden_admin" value="<%=Application("hidden_admin")%>">
隐身的管理员名字如：张三,李四<br>
<input type="text" name="aqjh_slbox" value="<%=Application("aqjh_slbox")%>">
可以查看私聊的人如：张三,李四<br>
<input type="text" name="aqjh_admin_send" value="<%=Application("aqjh_admin_send")%>">
财神爷，多个用|和|隔开，后面也要保留|<br>
<input type="text" name="aqjh_guibin" value="<%=Application("aqjh_guibin")%>">
江湖贵宾(贵宾不参与任何江湖恩怨)，多个用|和|隔开，后面也要保留|<br>
<input type="text" name="aqjh_adminuser" value="<%=Application("aqjh_adminuser")%>">
后台管理账号<br>
<input type="password" name="aqjh_adminkey" value="<%=Application("aqjh_adminkey")%>">
后台管理密码</p>
<p> 
<input type="submit" name="Submit" value="修改设置">
<input type="hidden" name="set" value="ok">
</p>
</form>
</body>
</html>
<%
response.end
else
call setconst()
end if
sub setconst()
%>
<!--#include file="chkform.asp"-->
<%
'*************************************************************
'*以下为江湖系统设置请按自己的需要设置，设置前最好备份一下原文件！*
'*************************************************************
writeStr=writeStr & "Application("&""""&"aqjh_chatbgcolor"&""""&")=" &"""" & request.form("aqjh_chatbgcolor")& """"&"      '聊天室背景颜色代码"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_chatimage"&""""&")=" &"""" & request.form("aqjh_chatimage")& """"&"      '聊天室外部边框图片"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_chatcolor"&""""&")=" &"""" & request.form("aqjh_chatcolor")& """"&"      '聊天室文字区背景色"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_iplocktime"&""""&")="& request.form("aqjh_iplocktime")&"      '锁IP时间"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_maxtimeout"&""""&")="& request.form("aqjh_maxtimeout")&"      '泡点超时时间，不要设置太长"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_maxnpc"&""""&")="& request.form("aqjh_maxnpc")&"      'npc总在线人数"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_npcwp"&""""&")="& request.form("aqjh_npcwp")&"      'npc暴涨比较值"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_npcff"&""""&")="& request.form("aqjh_npcff")&"      'npc发放需要的等级（管理级）"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_baowuyin"&""""&")="& request.form("aqjh_baowuyin")&"      '每一级修练所得到的银两"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_baowuxl"&""""&")="& request.form("aqjh_baowuxl")&"      '每次修练的级数，如100级的需修练4次！他所得到的上限就是4！这个值是用战斗等级取整的"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_killman"&""""&")="& request.form("aqjh_killman")&"      '江湖一天的杀人数，夺宝一次算一次杀人"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_wgsx="&request.form("aqjh_wgsx")&"      '每升一级的武功加值,此值当aqjh_sx=1时才有作用"&chr(13)&chr(10)
writeStr=writeStr & "aqjh_nlsx="&request.form("aqjh_nlsx")&"      '每升一级的内力加值,此值当aqjh_sx=1时才有作用"&chr(13)&chr(10)
writeStr=writeStr & "aqjh_tlsx="&request.form("aqjh_tlsx")&"      '每升一级的体力加值,此值当aqjh_sx=1时才有作用"&chr(13)&chr(10)
writeStr=writeStr & "aqjh_gjsx="&request.form("aqjh_gjsx")&"      '每升一级的攻击加值,此值当aqjh_sx=1时才有作用"&chr(13)&chr(10)
writeStr=writeStr & "aqjh_fysx="&request.form("aqjh_fysx")&"      '每升一级的防御加值,此值当aqjh_sx=1时才有作用"&chr(13)&chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_chat_maxpeople"&""""&")="& request.form("aqjh_chat_maxpeople")&"      '江湖最大人数限制"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_closedoor"&""""&")="& request.form("aqjh_closedoor")&"      '聊天室关门，1为关门"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_baowu"&""""&")="& request.form("aqjh_baowu")&"      '是否有宝物当为0时没有宝物！宝物出现是会影响聊天室速度的，服务器不好关闭！"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_sx="&request.form("aqjh_sx")&"      '是否有上限0没有上限，1有上限"&chr(13)&chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_automanname"&""""&")=" &"""" & request.form("aqjh_automanname")& """"&"      '聊神名称设置"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_userout"&""""&")=" &"""" & request.form("aqjh_userout")& """"&"      '退出时聊天室显示的信息"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_dieip=" &"""" & request.form("aqjh_dieip")& """"&"      'ip永久封锁使用前按格式书，多个用分号隔开，后面请保留有分号"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_baowuname"&""""&")=" &"""" & request.form("aqjh_baowuname")& """"&"      '宝物名及宝物说明"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_baowusm"&""""&")=" &"""" & request.form("aqjh_baowusm")& """"&"      '宝物在聊天室中显示信息"& chr(13) & chr(10)
writeStr=writeStr & "'***************************************************"& chr(13) & chr(10)
writeStr=writeStr & "'*以下江湖管理设置，切记不要改错，否则无法正常运行 *"& chr(13) & chr(10)
writeStr=writeStr & "'***************************************************"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_chatroomname"&""""&")=" &"""" & request.form("aqjh_chatroomname")& """"&"      '江湖名称"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_homepage"&""""&")=" &"""" & request.form("aqjh_homepage")& """"&"      '江湖主页"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_sn"&""""&")=" &"""" & request.form("aqjh_sn")& """"&"      '江湖序列号，正版标志，修改后后果自负"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_user"&""""&")=" &"""" & request.form("aqjh_user")& """"&"      '江湖站长"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_qq"&""""&")=" &"""" & request.form("aqjh_qq")& """"&"      '站长qq"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_email"&""""&")=" &"""" & request.form("aqjh_email")& """"&"      '站长email"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_admin"&""""&")=" &"""" & request.form("aqjh_admin")& """"&"      '设置10级聊管，多个用逗号隔开"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"hidden_admin"&""""&")=" &"""" & request.form("hidden_admin")& """"&"      '隐身人员，多个用逗号隔开"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_slbox"&""""&")=" &"""" & request.form("aqjh_slbox")& """"&"      '可以查看私聊的人，多个用逗号隔开"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_admin_send"&""""&")=" &"""" & request.form("aqjh_admin_send")& """"&"      '财神爷，多个用|和|隔开"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_guibin"&""""&")=" &"""" & request.form("aqjh_guibin")& """"&"      '江湖贵宾(贵宾不参与任何江湖恩怨)，多个用|和|隔开"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_adminuser"&""""&")=" &"""" & request.form("aqjh_adminuser")& """"&"      '站长管理用户名"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_adminkey"&""""&")=" &"""" & request.form("aqjh_adminkey")& """"&"      '站长管理密码"& chr(13) & chr(10)
writeStr=writeStr & "'*********************************************************"& chr(13) & chr(10)
writeStr=writeStr & "'*如果要设置广告内容请见：chat/jhchat.asp文件开始处"& chr(13) & chr(10)
writeStr=writeStr & "'*泡分的点数可以由自己设置，详见：paodian.asp文件开始处"& chr(13) & chr(10)
writeStr=writeStr & "'修改招式等级修改chat/sjfunc/czdj.asp"
%>
<%
tpstr=replace(writestr,chr(13) & chr(10) ,"<br>")
tpstr="&lt;%<br>" & tpstr & "<br>%&gt;"
writestr="<" & "%" & chr(13) & chr(10) & writestr & chr(13) & chr(10) & "%" & ">"
response.write tpstr
if IsObjInstalled("Scripting.FileSystemObject") then
toppath = Server.Mappath("../config.asp")
Set fs = CreateObject("scripting.filesystemobject")
If Not Fs.FILEEXISTS(toppath) Then 
Set Ts = fs.createtextfile(toppath, True)
Ts.close
end if
Set Ts= Fs.OpenTextFile(toppath,2)
Ts.writeline (writeStr)
Ts.Close
Set Ts=Nothing
Set Fs=Nothing
%>
<script>
alert("您的服务器支持FSO权限，新的config.asp已自动生成")
</script>
<%
else
%>
<script>
alert("您的服务器不支持FSO权限，请手工拷贝代码覆盖config.asp中原来的代码")
</script>
<%
end if
end sub
Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function
Function rrstr(str)
str="|" & str & "|"
str=replace(str,"||","|")
rrstr=str
End Function 
%>