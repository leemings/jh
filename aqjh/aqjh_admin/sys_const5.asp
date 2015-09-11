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
<!--#include file="../const5.asp"-->
<html>
<head>
<title>系统参数</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=css/css.css type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<%if not request.form("set")="ok" then%>
<div align="center"> 【论坛奖罚参数】</div>
<hr noshade size="1" color=009900>
<b>［提示］</b>注意，所有参数必填不可为空，不可以用全角字符输入参数。
<a href="javascript:history.go(0)">刷新</a> 
<hr noshade size="1" color=009900>
***论坛相关参数[1000豆点=1个金豆=2元金卡]因此不要设置太多***<br>
以下奖励和惩罚的均为泡豆点数
<form name="form1" method="post" action="">
<font color=red>************ 限制 ***********</font>
<BR>
<input type="text" name="bbs_add_dj" value="<%=bbs_add_dj%>">
允许发贴或回贴需要的江湖等级<br>
<input type="text" name="de_grade" value="<%=de_grade%>">
充许发表最低管理等级<br>
<font color=red>************ 增加 ***********</font>
<BR>
<input type="text" name="bbs_add1" value="<%=bbs_add1%>">
发贴+加<br>
<input type="text" name="bbs_add2" value="<%=bbs_add2%>">
回复+加<br>
<input type="text" name="bbs_add3" value="<%=bbs_add3%>">
拉前主题+加<br>
<input type="text" name="bbs_add4" value="<%=bbs_add4%>">
总固顶+加<br>
<input type="text" name="bbs_add5" value="<%=bbs_add5%>">
固顶+加<br>
<input type="text" name="bbs_add6" value="<%=bbs_add6%>">
精华区+加<br>
<font color=red>************ 扣除 ***********</font>
<BR>
<input type="text" name="bbs_del1" value="<%=bbs_del1%>">
删除主题-减<br>
<input type="text" name="bbs_del2" value="<%=bbs_del2%>">
删除回复-减<br>
<input type="text" name="bbs_del3" value="<%=bbs_del3%>">
关闭主题-减<br>
<input type="text" name="bbs_del4" value="<%=bbs_del4%>">
移动主题-减<br>
<font color=red>************ 系统 ***********</font>
<BR>
<input type="text" name="clubname" value="<%=clubname%>">
论坛名称<br>
<input type="text" name="allclass" value="<%=allclass%>">
打开社区首页自动展开所有论坛0=否 1=是<br>
<input type="text" name="badwords" value="<%=badwords%>">
过滤敏感字（多字请用“|”分隔）<br>
<input type="text" name="kickwords" value="<%=kickwords%>">
封锁用户（前,后,中间用“|”分隔）<br>
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
writeStr=writeStr & "'***论坛相关参数[1000豆点=1个金豆=2元金卡]因此不要设置太多***"& chr(13) & chr(10)
writeStr=writeStr & "'以下奖励和惩罚的均为泡豆点数"& chr(13) & chr(10)
writeStr=writeStr & "'-------------限制-----------------"& chr(13) & chr(10)
writeStr=writeStr & "bbs_add_dj="& request.form("bbs_add_dj")&"      '允许发贴或回贴需要的江湖等级"& chr(13)& chr(10)
writeStr=writeStr & "de_grade="& request.form("de_grade")&"      '充许发表最低管理等级"&chr(13) & chr(10)
writeStr=writeStr & "'-------------增加-----------------"& chr(13) & chr(10)
writeStr=writeStr & "bbs_add1="& request.form("bbs_add1")&"      '发贴+加"&chr(13) & chr(10)
writeStr=writeStr & "bbs_add2="& request.form("bbs_add2")&"      '回复+加"& chr(13) & chr(10)
writeStr=writeStr & "bbs_add3="& request.form("bbs_add3")&"      '拉前主题+加"& chr(13) & chr(10)
writeStr=writeStr & "bbs_add4="& request.form("bbs_add4")&"      '总固顶+加"& chr(13) & chr(10)
writeStr=writeStr & "bbs_add5="& request.form("bbs_add5")&"      '固顶+加"& chr(13) & chr(10)
writeStr=writeStr & "bbs_add6="& request.form("bbs_add6")&"      '精华区+加"& chr(13) & chr(10)
writeStr=writeStr & "'-------------扣除-----------------"& chr(13) & chr(10)
writeStr=writeStr & "bbs_del1="& request.form("bbs_del1")&"      '删除主题-减"& chr(13) & chr(10)
writeStr=writeStr & "bbs_del2="&""""& request.form("bbs_del2")&""""&"      '删除回复-减"& chr(13) & chr(10)
writeStr=writeStr & "bbs_del3="&""""& request.form("bbs_del3")&""""&"      '关闭主题-减"& chr(13) & chr(10)
writeStr=writeStr & "bbs_del4="&""""& request.form("bbs_del4")&""""&"      '移动主题-减"& chr(13) & chr(10)
writeStr=writeStr & "'-------------系统-----------------"& chr(13) & chr(10)
writeStr=writeStr & "clubname="&""""& request.form("clubname")&""""&"      '论坛名称"& chr(13) & chr(10)
writeStr=writeStr & "cluburl="&""""& request.form("cluburl")&""""&"      '论坛url地址"& chr(13) & chr(10)
writeStr=writeStr & "badwords="&""""& request.form("badwords")&""""&"      '过滤敏感字（多字请用“|”分隔）"& chr(13) & chr(10)
writeStr=writeStr & "allclass="& request.form("allclass")&"      '打开社区首页自动展开所有论坛0=否 1=是"& chr(13) & chr(10)
writeStr=writeStr & "kickwords="&""""& request.form("kickwords")&""""&"      '封锁用户（前,后,中间用“|”分隔）"& chr(13) & chr(10)
writeStr=writeStr & "'*********************************************************"
tpstr=replace(writestr,chr(13) & chr(10) ,"<br>")
tpstr="&lt;%<br>" & tpstr & "<br>%&gt;"
writestr="<" & "%" & chr(13) & chr(10) & writestr & chr(13) & chr(10) & "%" & ">"
response.write tpstr
if IsObjInstalled("Scripting.FileSystemObject") then
toppath = Server.Mappath("../const5.asp")
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
alert("您的服务器支持FSO权限，新的const5.asp已自动生成")
</script>
<%
else
%>
<script>
alert("您的服务器不支持FSO权限，请手工拷贝代码覆盖const5.asp中原来的代码")
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