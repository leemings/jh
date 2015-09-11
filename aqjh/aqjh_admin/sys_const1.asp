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
<!--#include file="../const1.asp"-->
<html>
<head>
<title>系统参数</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=css/css.css type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<%if not request.form("set")="ok" then%>
<div align="center"> 【新人状态配置】</div>
<hr noshade size="1" color=009900>
<b>［提示］</b>注意，所有参数必填不可为空，不可以用全角字符输入参数。
<a href="javascript:history.go(0)">刷新</a> 
<hr noshade size="1" color=009900>
<form name="form1" method="post" action="">
<input type="text" name="reg_ml" value="<%=reg_ml%>">
魅力<br>
<input type="text" name="reg_dd" value="<%=reg_dd%>">
道德<br>
<input type="text" name="reg_wg" value="<%=reg_wg%>">
武功<br>
<input type="text" name="reg_nl" value="<%=reg_nl%>">
内力<br>
<input type="text" name="reg_tl" value="<%=reg_tl%>">
体力<br>
<input type="text" name="reg_gj" value="<%=reg_gj%>">
攻击<br>
<input type="text" name="reg_fy" value="<%=reg_fy%>">
防御<br>
<input type="text" name="reg_fl" value="<%=reg_fl%>">
法力<br>
<input type="text" name="reg_zz" value="<%=reg_zz%>">
知质<br>
<input type="text" name="reg_qg" value="<%=reg_qg%>">
轻功<br>
<input type="text" name="reg_yl" value="<%=reg_yl%>">
银两<br>
<input type="text" name="reg_ck" value="<%=reg_ck%>">
存款<br>
<input type="text" name="reg_jb" value="<%=reg_jb%>">
金币<br>
<input type="text" name="reg_dj" value="<%=reg_dj%>">
注册时的等级<font color=red>[建议使用默认值15级，因为系统设置的18级以下为新手]</font><br>
<input type="text" name="reg_hyjk" value="<%=reg_hyjk%>">
会员金卡<br>
<input type="text" name="reg_hydj" value="<%=reg_hydj%>">
会员等级[0至4级，0为非等级会员]<br>
<input type="text" name="reg_hydate" value="<%=reg_hydate%>">
会员期限[30为一个月的时间]<br>
<input type="text" name="reg_hycard" value="<%=reg_hycard%>">
卡片[正确格式(红色部分)：<font color=red><u>卡片名1|数量;</u>卡片名2|数量;</font>]<br><br>
<input type="text" name="reg_yp" value="<%=reg_yp%>">
药品[正确格式(红色部分)：<font color=red><u>药品名1|数量;</u>药品名2|数量;</font>]<br><br>
<input type="text" name="reg_xh" value="<%=reg_xh%>">
鲜花[正确格式(红色部分)：<font color=red><u>鲜花名1|数量;</u>鲜花名2|数量;</font>]<br><br>
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
writeStr=writeStr & "'*****************新人设置*********************"& chr(13) & chr(10)
writeStr=writeStr & "reg_ml="& request.form("reg_ml")&"      '魅力"& chr(13)& chr(10)
writeStr=writeStr & "reg_dd="& request.form("reg_dd")&"      '道德"&chr(13) & chr(10)
writeStr=writeStr & "reg_wg="& request.form("reg_wg")&"      '武功"&chr(13) & chr(10)
writeStr=writeStr & "reg_nl="& request.form("reg_nl")&"      '内力"& chr(13) & chr(10)
writeStr=writeStr & "reg_tl="& request.form("reg_tl")&"      '内力"& chr(13) & chr(10)
writeStr=writeStr & "reg_gj="& request.form("reg_gj")&"      '攻击"& chr(13) & chr(10)
writeStr=writeStr & "reg_fy="& request.form("reg_fy")&"      '防御"& chr(13) & chr(10)
writeStr=writeStr & "reg_fl="& request.form("reg_fl")&"      '法力"& chr(13) & chr(10)
writeStr=writeStr & "reg_zz="& request.form("reg_zz")&"      '知质"& chr(13) & chr(10)
writeStr=writeStr & "reg_qg="& request.form("reg_qg")&"      '轻功"& chr(13) & chr(10)
writeStr=writeStr & "reg_yl="& request.form("reg_yl")&"      '银两"& chr(13) & chr(10)
writeStr=writeStr & "reg_ck="& request.form("reg_ck")&"      '存款"& chr(13) & chr(10)
writeStr=writeStr & "reg_jb="& request.form("reg_jb")&"      '金币"& chr(13) & chr(10)
writeStr=writeStr & "reg_dj="& request.form("reg_dj")&"      '注册等级"& chr(13) & chr(10)
writeStr=writeStr & "reg_hyjk="& request.form("reg_hyjk")&"      '会员金卡"& chr(13) & chr(10)
writeStr=writeStr & "reg_hydj="& request.form("reg_hydj")&"      '会员等级"& chr(13) & chr(10)
writeStr=writeStr & "reg_hydate="& request.form("reg_hydate")&"      '会员期限"& chr(13) & chr(10)
writeStr=writeStr & "reg_hycard="&""""& request.form("reg_hycard")&""""&"      '卡片"& chr(13) & chr(10)
writeStr=writeStr & "reg_yp="&""""& request.form("reg_yp")&""""&"      '药品"& chr(13) & chr(10)
writeStr=writeStr & "reg_xh="&""""& request.form("reg_xh")&""""&"      '鲜花"& chr(13) & chr(10)
writeStr=writeStr & "'*********************************************************"
%>
<%
tpstr=replace(writestr,chr(13) & chr(10) ,"<br>")
tpstr="&lt;%<br>" & tpstr & "<br>%&gt;"
writestr="<" & "%" & chr(13) & chr(10) & writestr & chr(13) & chr(10) & "%" & ">"
response.write tpstr
if IsObjInstalled("Scripting.FileSystemObject") then
toppath = Server.Mappath("../const1.asp")
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
alert("您的服务器支持FSO权限，新的const1.asp已自动生成")
</script>
<%
else
%>
<script>
alert("您的服务器不支持FSO权限，请手工拷贝代码覆盖const1.asp中原来的代码")
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