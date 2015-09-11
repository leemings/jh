<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
if aqjh_name<>Application("aqjh_user") then Response.Redirect "../error.asp?id=486"
%>
<!--#include file="../const3.asp"-->
<html>
<head>
<title>系统参数</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=css/css.css type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<%if not request.form("set")="ok" then%>
<div align="center"> 【开关配置】</div>
<hr noshade size="1" color=009900>
<b>［提示］</b>注意，所有参数必填不可为空，不可以用全角字符输入参数。
<a href="javascript:history.go(0)">刷新</a> 
<hr noshade size="1" color=009900>
<form name="form1" method="post" action="">
<input type="text" name="aqjh_ifnpc" value="<%=aqjh_ifnpc%>">
NPC开关,1为开,0为关<br>
<input type="text" name="aqjh_chatkptime" value="<%=aqjh_chatkptime%>">
聊天大厅用卡时间限制，每小时前xx分钟，不能超过60<br>
<input type="text" name="aqjh_onlinekill" value="<%=aqjh_onlinekill%>">
聊天大厅动武人限，低于此人数，禁止动武<br>
<input type="text" name="chatshuaxin_time" value="<%=chatshuaxin_time%>">
聊天室防刷新时间为xx秒<br>
<input type="text" name="aqjh_dongtai" value="<%=aqjh_dongtai%>">
是否显示动态广告，具体见：chat/gg.js文件内容！0为不显示，1为显示<br>
<input type="text" name="aqjh_disproxy" value="<%=aqjh_disproxy%>">
代理服务器，1为限制，0为不限制<br>
<input type="text" name="aqjh_yjdh" value="<%=aqjh_yjdh%>">
是否限制一机多号,1为限制,0为不限制<br>
<input type="text" name="aqjh_myie" value="<%=aqjh_myie%>">
是否限制myie多窗口浏览器,1为限制,0为不限制<br>
<input type="text" name="aqjh_sfkf" value="<%=aqjh_sfkf%>">
是否开放掌门创建房间功能，0为开放，1为关闭<br>
<input type="text" name="aqjh_chat1" value="<%=aqjh_chat1%>">
恩怨房间名称<br>
<input type="text" name="aqjh_chat2" value="<%=aqjh_chat2%>">
决战房间名称<br>
<input type="text" name="aqjh_chat3" value="<%=aqjh_chat3%>">
npc房间名称，默认为聊天大厅<br>
<br><br>
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
writeStr=writeStr & "'*****************开关配置*********************"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_ifnpc="& request.form("aqjh_ifnpc")&"      'NPC开关,1为开,0为关"& chr(13)& chr(10)
writeStr=writeStr & "aqjh_chatkptime="& request.form("aqjh_chatkptime")&"      '聊天大厅用卡时间限制，每小时前xx分钟，不能超过60"&chr(13) & chr(10)
writeStr=writeStr & "aqjh_onlinekill="& request.form("aqjh_onlinekill")&"      '聊天大厅动武人限，低于此人数，禁止动武"&chr(13) & chr(10)
writeStr=writeStr & "chatshuaxin_time="& request.form("chatshuaxin_time")&"      '聊天室防刷新时间为xx秒"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_dongtai="& request.form("aqjh_dongtai")&"      '是否显示动态广告，具体见：chat/gg.js文件内容！0为不显示，1为显示"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_disproxy="& request.form("aqjh_disproxy")&"      '代理服务器，1为限制，0为不限制"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_yjdh="& request.form("aqjh_yjdh")&"      '是否限制一机多号,1为限制,0为不限制"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_myie="& request.form("aqjh_myie")&"      '是否限制myie登陆,1为限制,0为不限制"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_sfkf="& request.form("aqjh_sfkf")&"      '是否开放掌门创建房间功能，0为开放，1为关闭"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_chat1="&""""& request.form("aqjh_chat1")&""""&"      '恩怨房间名称"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_chat2="&""""& request.form("aqjh_chat2")&""""&"      '决战房间名称"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_chat3="&""""& request.form("aqjh_chat3")&""""&"      'npc房间名称，默认为聊天大厅"& chr(13) & chr(10)
writeStr=writeStr & "'*********************************************************"
%>
<%
tpstr=replace(writestr,chr(13) & chr(10) ,"<br>")
tpstr="&lt;%<br>" & tpstr & "<br>%&gt;"
writestr="<" & "%" & chr(13) & chr(10) & writestr & chr(13) & chr(10) & "%" & ">"
response.write tpstr
if IsObjInstalled("Scripting.FileSystemObject") then
toppath = Server.Mappath("../const3.asp")
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
alert("您的服务器支持FSO权限，新的const3.asp已自动生成")
</script>
<%
else
%>
<script>
alert("您的服务器不支持FSO权限，请手工拷贝代码覆盖const3.asp中原来的代码")
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