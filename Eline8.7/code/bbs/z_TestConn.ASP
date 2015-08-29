<%
	dim aconn,rUrl
	dim aconnstr
	const TestDb="data/e_testchen.asp"		'修改数据库路径和数据库名称（相对路径）
	Set aconn = Server.CreateObject("ADODB.Connection")
	aconnstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(TestDb)
	'如果你的服务器采用较老版本Access驱动，请用下面连接方法
	'aconnstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(TestDb)
	
	aconn.Open aconnstr


'==========通用函数========================
function GetLbName(Lid)
	dim grs
	set grs=aconn.execute("Select lb From [testlb] where id="&Lid)
	if grs.eof then
		GetLbName="<font color=red>已删除</font>"
	else
		GetLbName=grs(0)
	end if 
end function
'======================
sub suc()
%>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:75%">
	<tr><th height=25>开心辞典 成功信息</th></tr>
	<tr><td width="100%" class=tablebody1><b>操作成功：</b><br><br><li>欢迎光临<%=Forum_info(0)%>开心辞典
	<font color=navy><%=sucmsg%></font><br></td></tr>
	<tr><td align=center height=26 class=tablebody2><a href=<%if rUrl="" or isnull(rUrl) then%><%=Request.ServerVariables("HTTP_REFERER")%><%else%><%=rUrl%><%end if%>>返&nbsp;&nbsp;回</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=z_test.asp>开心辞典</a></td></tr>
</table>
<%
end sub

sub test_err() 
%>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:75%">
	<tr><th height=25>开心辞典 错误信息</th></tr>
	<tr><td width="100%" class=tablebody1><b>产生错误的可能原因：</b><br><br><li>您是否仔细阅读了<a href="Z_TestHelp.asp">帮助文件</a>，可能您还没有登陆或者不具有使用当前功能的权限
	<font color=<%=Forum_body(8)%>><%=errmsg%></font><br></td></tr>
	<tr><td align=center height=26 class=tablebody2><a href="javascript:history.go(-1)">返回上一页</a></td></tr>
</table>
<%
end sub

sub Admin_Head()
%>
<br>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
	<tr>
		<td align=left valign=middle class=TopLighNav1> &nbsp;<a href="Z_test.asp" title="返回开心辞典首页">开心辞典</a> | <a href="z_testSetting.asp" title="辞典参数设置、辞典类别管理">辞典管理</a> | <a href="Z_TestAdminUpLoad.asp" title="审核上传的题目">上传管理</a> | <a href="Z_TestEdit.asp" title="管理已有题库">题库管理</a>  | <a href="Z_TestAdminUser.asp" title="管理辞典用户">用户管理</a> | <a href="Z_TestLayIn.asp" title="题库批量导入">题库批量导入</a> | <a href=javascript:history.go(-1)>返回上一页</a></td>
	</tr> 
</table>
<br>
<%
end sub

sub User_Head()
%>
<br>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
	<tr>
		<td align=left valign=middle class=TopLighNav1> &nbsp;<a href="Z_test.asp">开心辞典</a> | <a href="Z_test.asp?action=PaiXing">答题龙虎榜</a> | <span lang="zh-cn"><a href="Z_TestUpload.ASP">上载选择题</a></span> | <span lang="zh-cn"><a href="Z_TestUpload.ASP?action=addtk">上载填空题</a></span> | <a href=javascript:history.go(-1)>返回上一页</a></td>
	</tr> 
</table>
<br>
<%
end sub
'===================
sub NewsLog(lclass,content,userN)
'我来了 2002.12.01 写入事件
	if UserID="" or (not isnumeric(UserID)) then UserID=0 
	if content="" then content="未知的操作"
	aconn.execute("insert into testnews ([class],UserName,UserID,content,AddTime,IP) values('"&lclass&"','"&userN&"',"&UserID&",'"&content&"',now(),'"&Request.ServerVariables("REMOTE_ADDR")&"')")
end sub
%>