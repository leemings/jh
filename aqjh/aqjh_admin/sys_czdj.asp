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
<!--#include file="../chat/sjfunc/czdj.asp"-->
<html>
<head>
<title>系统参数</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=css/css.css type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<%if not request.form("set")="ok" then%>
<div align="center"> 【操作限制设置】</div>
<hr noshade size="1" color=009900>
<b>［提示］</b>注意，所有参数必填不可为空，不可以用全角字符输入参数。
<a href="javascript:history.go(0)">刷新</a> 
<hr noshade size="1" color=009900>
<form name="form1" method="post" action="">
*****************以下需要的是战斗等级*****************<Br>
<input type="text" name="jhdj_qlcy" value="<%=jhdj_qlcy%>">
千里传音等级<br>
<input type="text" name="jhdj_duyao" value="<%=jhdj_duyao%>">
下毒等级<br>
<input type="text" name="jhdj_tq" value="<%=jhdj_tq%>">
偷钱等级<br>
<input type="text" name="jhdj_xx" value="<%=jhdj_xx%>">
吸星大法<br>
<input type="text" name="jhdj_aq" value="<%=jhdj_aq%>">
暗器等级<br>
<input type="text" name="jhdj_ai" value="<%=jhdj_ai%>">
魔宝等级<br>
<input type="text" name="jhdj_fz" value="<%=jhdj_fz%>">
发招攻击<br>
<input type="text" name="jhdj_cs" value="<%=jhdj_cs%>">
传授内力<br>
<input type="text" name="jhdj_sq" value="<%=jhdj_sq%>">
送钱等级<br>
<input type="text" name="jhdj_so" value="<%=jhdj_so%>">
送法力等级<br>
<input type="text" name="jhdj_zs" value="<%=jhdj_zs%>">
赠送和心动等级<br>
<input type="text" name="jhdj_pm" value="<%=jhdj_pm%>">
拍卖等级<br>
<input type="text" name="jhdj_bs" value="<%=jhdj_bs%>">
拜师等级<br>
<input type="text" name="jhdj_jj" value="<%=jhdj_jj%>">
传经验的等级<br>
<input type="text" name="jhdj_zz" value="<%=jhdj_zz%>">
转账需要的等级<br>
<input type="text" name="jhdj_nh" value="<%=jhdj_nh%>">
怒吼需要等级<br>
<input type="text" name="jhdj_xt" value="<%=jhdj_xt%>">
心跳需要等级<br>
<input type="text" name="jhdj_hy" value="<%=jhdj_hy%>">
在线求婚<br>
<input type="text" name="jhdj_db" value="<%=jhdj_db%>">
赌博作庄<br>
<input type="text" name="jhdj_qr" value="<%=jhdj_qr%>">
情人、飞吻需要等级<br>
<input type="text" name="jhdj_dt" value="<%=jhdj_dt%>">
单挑需要等级<br>
<input type="text" name="jhdj_chi" value="<%=jhdj_chi%>">
求签等级<br>
<input type="text" name="jhdj_hua" value="<%=jhdj_hua%>">
送花等级<br>
<input type="text" name="jhdj_zl" value="<%=jhdj_zl%>">
暂离开等级<br>
<input type="text" name="jhdj_xq" value="<%=jhdj_xq%>">
心情功能等级<br>
<input type="text" name="jhdj_ti" value="<%=jhdj_ti%>">
标题等级<br>
<input type="text" name="jhdj_nj" value="<%=jhdj_nj%>">
念经需要等级<br>
<input type="text" name="jhdj_jhtw" value="<%=jhdj_jhtw%>">
加入天网等级<br>
<input type="text" name="jhdj_frag" value="<%=jhdj_frag%>">
倒夜香等级<br>
<input type="text" name="aqjh_jhcz" value="<%=aqjh_jhcz%>">
出家等级<br>
*****************以下需要的是管理等级*****************<Br>
<input type="text" name="grade_dx" value="<%=grade_dx%>">
点穴操作等级<br>
<input type="text" name="grade_jx" value="<%=grade_jx%>">
解穴操作等级<br>
<input type="text" name="grade_db" value="<%=grade_db%>">
逮捕操作等级<br>
<input type="text" name="grade_zl" value="<%=grade_zl%>">
坐牢操作等级<br>
<input type="text" name="grade_jj" value="<%=grade_jj%>">
监禁操作等级<br>
<input type="text" name="grade_jc" value="<%=grade_jc%>">
解除监禁等级<br>
<input type="text" name="grade_zs" value="<%=grade_zs%>">
斩首操作等级<br>
<input type="text" name="grade_jg" value="<%=grade_jg%>">
警告操作等级<br>
<input type="text" name="grade_see" value="<%=grade_see%>">
查看状态<br>
<input type="text" name="grade_cf" value="<%=grade_cf%>">
册封等级固定<br>
<input type="text" name="grade_gg" value="<%=grade_gg%>">
发表公告的等级<br>
<input type="text" name="grade_tr" value="<%=grade_tr%>">
踢人等级<br>
<input type="text" name="grade_fd" value="<%=grade_fd%>">
放大文字<br>
<input type="text" name="grade_ip" value="<%=grade_ip%>">
查ip等级<br>
<input type="text" name="grade_bo" value="<%=grade_bo%>">
轰炸<br>
<input type="text" name="grade_jr" value="<%=grade_jr%>">
今日话题<br>
<input type="text" name="grade_tjf" value="<%=grade_tjf%>">
通缉犯等级<br>
**********点歌控制****************************************<br>
<input type="text" name="pyyl" value="<%=pyyl%>">
点给朋友需要的金币数<br>
<input type="text" name="pyyls" value="<%=pyyls%>">
点给朋友需要的银两数<br>
<input type="text" name="xyyl" value="<%=xyyl%>">
点歌给大家需要金币数<br>
<input type="text" name="xyyls" value="<%=xyyls%>">
点歌给大家需要银两数<br>
<input type="text" name="xzsj" value="<%=xzsj%>">
是否有时间限制(true为有,false为无)<br>
<input type="text" name="xzsjs" value="<%=xzsjs%>">
点歌间隔时间限制<br>
<br>
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
writeStr=writeStr & "'*********以下需要的是战斗等级*************************"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_qlcy="& request.form("jhdj_qlcy")&"      '千里传音等级"& chr(13)& chr(10)
writeStr=writeStr & "jhdj_duyao="& request.form("jhdj_duyao")&"      '下毒等级"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_tq="& request.form("jhdj_tq")&"      '偷钱等级"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_xx="& request.form("jhdj_xx")&"      '吸星大法"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_aq="& request.form("jhdj_aq")&"      '暗器等级"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_ai="& request.form("jhdj_ai")&"      '魔宝等级"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_fz="& request.form("jhdj_fz")&"      '发招攻击"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_cs="& request.form("jhdj_cs")&"      '传授内力"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_sq="& request.form("jhdj_sq")&"      '送钱等级"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_so="& request.form("jhdj_so")&"      '送法力等级"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_zs="& request.form("jhdj_zs")&"      '赠送和心动等级"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_pm="& request.form("jhdj_pm")&"      '拍卖等级"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_bs="& request.form("jhdj_bs")&"      '拜师等级"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_jj="& request.form("jhdj_jj")&"      '传经验的等级"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_zz="& request.form("jhdj_zz")&"      '转账需要的等级"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_nh="& request.form("jhdj_nh")&"      '怒吼需要等级"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_xt="& request.form("jhdj_xt")&"      '心跳需要等级"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_hy="& request.form("jhdj_hy")&"      '在线求婚"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_db="& request.form("jhdj_db")&"      '赌博作庄"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_qr="& request.form("jhdj_qr")&"      '情人、飞吻需要等级"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_dt="& request.form("jhdj_dt")&"      '单挑需要等级"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_chi="& request.form("jhdj_chi")&"      '求签等级"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_hua="& request.form("jhdj_hua")&"      '送花等级"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_zl="& request.form("jhdj_zl")&"      '暂离开等级"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_xq="& request.form("jhdj_xq")&"      '心情功能等级"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_ti="& request.form("jhdj_ti")&"      '标题等级"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_nj="& request.form("jhdj_nj")&"      '念经需要等级"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_jhtw="& request.form("jhdj_jhtw")&"      '加入天网等级"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_frag="& request.form("jhdj_frag")&"      '倒夜香等级"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_jhcz="& request.form("aqjh_jhcz")&"      '出家等级"& chr(13) & chr(10)
writeStr=writeStr & "'*********以下需要的是管理等级*************************"& chr(13) & chr(10)
writeStr=writeStr & "grade_dx="& request.form("grade_dx")&"      '点穴操作等级"& chr(13) & chr(10)
writeStr=writeStr & "grade_jx="& request.form("grade_jx")&"      '解穴操作等级"& chr(13) & chr(10)
writeStr=writeStr & "grade_db="& request.form("grade_db")&"      '逮捕操作等级"& chr(13) & chr(10)
writeStr=writeStr & "grade_zl="& request.form("grade_zl")&"      '坐牢操作等级"& chr(13) & chr(10)
writeStr=writeStr & "grade_jj="& request.form("grade_jj")&"      '监禁操作等级"& chr(13) & chr(10)
writeStr=writeStr & "grade_jc="& request.form("grade_jc")&"      '解除监禁等级"& chr(13) & chr(10)
writeStr=writeStr & "grade_zs="& request.form("grade_zs")&"      '斩首操作等级"& chr(13) & chr(10)
writeStr=writeStr & "grade_jg="& request.form("grade_jg")&"      '警告操作等级"& chr(13) & chr(10)
writeStr=writeStr & "grade_see="& request.form("grade_see")&"      '查看状态"& chr(13) & chr(10)
writeStr=writeStr & "grade_cf="& request.form("grade_cf")&"      '册封等级固定"& chr(13) & chr(10)
writeStr=writeStr & "grade_gg="& request.form("grade_gg")&"      '发表公告的等级"& chr(13) & chr(10)
writeStr=writeStr & "grade_tr="& request.form("grade_tr")&"      '踢人等级"& chr(13) & chr(10)
writeStr=writeStr & "grade_fd="& request.form("grade_fd")&"      '放大文字"& chr(13) & chr(10)
writeStr=writeStr & "grade_ip="& request.form("grade_ip")&"      '查ip等级"& chr(13) & chr(10)
writeStr=writeStr & "grade_bo="& request.form("grade_bo")&"      '轰炸"& chr(13) & chr(10)
writeStr=writeStr & "grade_jr="& request.form("grade_jr")&"      '今日话题"& chr(13) & chr(10)
writeStr=writeStr & "grade_tjf="& request.form("grade_tjf")&"      '通缉犯等级"& chr(13) & chr(10)
writeStr=writeStr & "'*********点歌控制***********************************"& chr(13) & chr(10)
writeStr=writeStr & "pyyl="& request.form("pyyl")&"      '点给朋友需要的金币数"& chr(13) & chr(10)
writeStr=writeStr & "pyyls="& request.form("pyyls")&"      '点给朋友需要的银两数"& chr(13) & chr(10)
writeStr=writeStr & "xyyl="& request.form("xyyl")&"      '点歌给大家需要金币数"& chr(13) & chr(10)
writeStr=writeStr & "xyyls="& request.form("xyyls")&"      '点歌给大家需要银两数"& chr(13) & chr(10)
writeStr=writeStr & "xzsj="& request.form("xzsj")&"      '是否有时间限制"& chr(13) & chr(10)
writeStr=writeStr & "xzsjs="& request.form("xzsjs")&"      '点歌间隔时间限制"
%>
<%
tpstr=replace(writestr,chr(13) & chr(10) ,"<br>")
tpstr="&lt;%<br>" & tpstr & "<br>%&gt;"
writestr="<" & "%" & chr(13) & chr(10) & writestr & chr(13) & chr(10) & "%" & ">"
response.write tpstr
if IsObjInstalled("Scripting.FileSystemObject") then
toppath = Server.Mappath("../chat/sjfunc/czdj.asp")
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
alert("您的服务器支持FSO权限，新的czdj.asp已自动生成")
</script>
<%
else
%>
<script>
alert("您的服务器不支持FSO权限，请手工拷贝代码覆盖czdj.asp中原来的代码")
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