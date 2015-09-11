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
<!--#include file="../const2.asp"-->
<html>
<head>
<title>系统参数</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=css/css.css type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<%if not request.form("set")="ok" then%>
<div align="center"> 【其它参数配置】</div>
<hr noshade size="1" color=009900>
<b>［提示］</b>注意，所有参数必填不可为空，不可以用全角字符输入参数。
<a href="javascript:history.go(0)">刷新</a> 
<hr noshade size="1" color=009900>
<form name="form1" method="post" action="">
<p><font color=red>*************每月月奖励设置***********</font><br>
<input type="text" name="jl_dd" value="<%=jl_dd%>">
每月在哪一天进行奖励(当前设为每月28号)<br>
<input type="text" name="jl_allvalue" value="<%=jl_allvalue%>">
被拉的人要达到的积分<br>
<input type="text" name="jl_yjf_top" value="<%=jl_yjf_top%>">
月积分奖励前多少名<br>
<input type="text" name="jl_yjf_jk" value="<%=jl_yjf_jk%>">
月积分奖励金卡系数<br>
<input type="text" name="jl_lr_jk" value="<%=jl_lr_jk%>">
拉人奖励金卡系数<br>
<input type="text" name="jl_lr_yl" value="<%=jl_lr_yl%>">
拉人奖励银两系数<br>
<p><font color=red>************禁止注册或登录的字眼***********</font><br>
<textarea name="aqjh_disloginname" cols="80" rows="10"><%=aqjh_disloginname%></textarea><br>
注意：别把自己的站长名添进去了，否则站长就进不了啦，一定要按照格式，如果不懂，请不要修改!<br>
<p><font color=red>************聊天室禁止的脏话***********</font><br>
<textarea name="aqjh_badword" cols="80" rows="10"><%=aqjh_badword%></textarea><br>
注意：一定要按照格式，如果不懂，请不要修改!<br>
<p><font color=red>************文字秀位于chat/picwords文件夹下***********</font><br>
<textarea name="aqjh_Zshow" cols="80" rows="10" readonly><%=aqjh_Zshow%></textarea><br>
注意：一定要按照格式，如果不懂，请不要修改!<br>
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
writeStr=writeStr & "'****************每月月奖励设置***************"& chr(13) & chr(10)
writeStr=writeStr & "jl_dd="&request.form("jl_dd")&"      '每月在哪一天进行奖励(当前设为每月28号)"&chr(13)&chr(10)
writeStr=writeStr & "jl_allvalue="&request.form("jl_allvalue")&"      '拉人需要达到一定的积分"&chr(13)&chr(10)
writeStr=writeStr & "jl_yjf_top="&request.form("jl_yjf_top")&"      '月积分奖励前多少名"&chr(13)&chr(10)
writeStr=writeStr & "jl_yjf_jk="&request.form("jl_yjf_jk")&"      '月积分奖励金卡系数"&chr(13)&chr(10)
writeStr=writeStr & "jl_lr_jk="&request.form("jl_lr_jk")&"      '拉人奖励金卡系数"&chr(13)&chr(10)
writeStr=writeStr & "jl_lr_yl="&request.form("jl_lr_yl")&"      '拉人奖励银两系数"&chr(13)&chr(10)
writeStr=writeStr & "'****************禁止注册或登录的字眼***************"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_disloginname="&""""& request.form("aqjh_disloginname")&""""&"      '禁止登录的字眼"& chr(13) & chr(10)
writeStr=writeStr & "'****************聊天室禁止的脏话***************"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_badword="&""""& request.form("aqjh_badword")&""""&"      '聊天室禁止的脏话"& chr(13) & chr(10)
writeStr=writeStr & "'****************文字秀位于chat/picwords文件夹下***************"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_Zshow="&""""& request.form("aqjh_Zshow")&""""&"      '秀字"& chr(13) & chr(10)
tpstr=replace(writestr,chr(13) & chr(10) ,"<br>")
tpstr="&lt;%<br>" & tpstr & "<br>%&gt;"
writestr="<" & "%" & chr(13) & chr(10) & writestr & chr(13) & chr(10) & "%" & ">"
response.write tpstr
if IsObjInstalled("Scripting.FileSystemObject") then
toppath = Server.Mappath("../const2.asp")
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
alert("您的服务器支持FSO权限，新的const2已自动生成")
</script>
<%
else
%>
<script>
alert("您的服务器不支持FSO权限，请手工拷贝代码覆盖const2中原来的代码")
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