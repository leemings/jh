<!-- #include file="setup.asp" -->
<!-- #include file="inc/MD5.asp" --><%

if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")

if membercode < 5 then error("<li>您没有权限进入后台")

if Request("menu")="out" then
session("pass")=""
response.redirect ""&Request.ServerVariables("script_name")&"/../"
end if

%>
<title><%=clubname%>后台管理 - Powered By BBSxp</title>
<%
select case Request("menu")
case ""
index
case "left"
left
case "top"
top2
case "login"
login

case "pass"
pass



end select


sub top2
%><body topmargin="0" rightmargin="0" leftmargin="0"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
	<tr>
		<td width="100%" class="a1" height="20" align="center"><b>
		<font face="幼圆">第五代 BBS 系统 -- BBSxp</font> <font color="ffff66">安全</font>
		<font color="ff0033">快速</font> <font color="33ff33">方便 </font>
		<font color="0000ff">可靠 </font><font color="000000">可升级</font></b> </td>
	</tr>
</table>
<%
end sub




sub pass

session("pass")=md5(""&Request("pass")&"")
if adminpassword<>session("pass") then error2("社区管理密码错误")


log("登录后台管理")


Dim theInstalledObjects(18)
theInstalledObjects(0) = "MSWC.AdRotator"
theInstalledObjects(1) = "MSWC.BrowserType"
theInstalledObjects(2) = "MSWC.NextLink"
theInstalledObjects(3) = "MSWC.Tools"
theInstalledObjects(4) = "MSWC.Status"
theInstalledObjects(5) = "MSWC.Counters"
theInstalledObjects(6) = "IISSample.ContentRotator"
theInstalledObjects(7) = "IISSample.PageCounter"
theInstalledObjects(8) = "MSWC.PermissionChecker"
theInstalledObjects(9) = "Scripting.FileSystemObject"
theInstalledObjects(10) = "adodb.connection"
theInstalledObjects(11) = "SoftArtisans.FileUp"
theInstalledObjects(12) = "SoftArtisans.FileManager"
theInstalledObjects(13) = "JMail.Message"
theInstalledObjects(14) = "CDONTS.NewMail"
theInstalledObjects(15) = "Persits.MailSender"
theInstalledObjects(16) = "LyfUpload.UploadFile"
theInstalledObjects(17) = "Persits.Upload.1"
theInstalledObjects(18) = "w3.upload"




%>
<table cellpadding="2" cellspacing="1" border="0" width="95%" align="center" class="a2">
	<tr>
		<td class="a1" colspan="2" height="25" align="center"><b>论坛信息统计</b></td>
	</tr>
	<tr class="a4">
		<td height="23" colspan="2">系统信息：论坛数 <b><%=conn.execute("Select count(id)from [bbsconfig]")(0)%></b> 
		主题数 <b><%=Application(CacheName&"CountForum")%></b> 用户数 <b><%=Application(CacheName&"CountUser")%></b>
		</td>
	</tr>
	<tr class="a4">
		<td width="50%" height="23">服务器域名：<%=Request.ServerVariables("server_name")%> 
		/ <%=Request.ServerVariables("LOCAL_ADDR")%></td>
		<td width="50%">脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
	</tr>
	<tr class="a4">
		<td width="50%" height="23">服务器软件的名称：<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
		<td width="50%">ACCESS 数据库路径：<a target="_blank" href="<%=datapath%><%=datafile%>"><%=datapath%><%=datafile%></a></td>
	</tr>
	<tr class="a4">
		<td width="50%" height="23">FSO 文本读写：<%If Not IsObjInstalled(theInstalledObjects(9)) Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%></td>
		<td width="50%">SA-FileUp 文件上传：<%If Not IsObjInstalled(theInstalledObjects(11)) Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%></td>
	</tr>
	<tr class="a4">
		<td width="50%" height="23">Jmail 组件支持：<%If Not IsObjInstalled(theInstalledObjects(13)) Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%></td>
		<td width="50%">CDONTS 组件支持：<%If Not IsObjInstalled(theInstalledObjects(14)) Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%></td>
	</tr>
</table>
<br>
<br>
<table cellpadding="2" cellspacing="1" border="0" width="95%" align="center" class="a2">
	<tr>
		<td class="a1" colspan="2" height="25" align="center"><b>管理快捷方式</b></td>
	</tr>
	<tr class="a4">
		<td width="20%" height="23">快速查找用户</td>
		<td width="80%" height="23">
		<form method="post" action="admin_user.asp?menu=user2">
			<input size="13" name="username"> <input type="submit" value="立刻查找">
		</td>
		</form>	</tr>
	<tr class="a4">
		<td width="20%" height="23">快捷功能链接</td>
		<td width="80%" height="23"><a href="admin_bbs.asp?menu=classs">
		<font color="000000">建立论坛数据</font></a> |
		<a href="admin_bbs.asp?menu=bbsmanage"><font color="000000">管理论坛资料</font></a> 
		| <a href="admin_bbs.asp?menu=upclubconfigok"><font color="000000">更新论坛资料</font></a></td>
	</tr>
</table>
<%
conn.execute("delete from [log] where logtime<"&SqlNowString&"-7")

htmlend
end sub




sub login





%> <br>
<br>
<br>
<form action="admin.asp" method="post">
	<input type="hidden" value="pass" name="menu">
	<table width="333" border="0" cellspacing="1" cellpadding="2" align="center" class="a2">
		<tr>
			<td width="328" class="a1" height="25">
			<div align="center">
				登录社区管理</div>
			</td>
		</tr>
		<tr>
			<td height="19" width="328" valign="top" class="a3">
			<div align="center">
				请输入管理密码: <input size="15" name="pass" type="password"><br>
				<input type="submit" value=" 登录 " name="Submit1">　<input type="reset" value=" 取消 " name="Submit">
			</div>
			</td>
		</tr>
	</table>
</form>
<%
htmlend
end sub

sub left

%>
<style>BODY {
	BACKGROUND: #799ae1; MARGIN: 0px;
}
.sec_menu {
	BORDER-RIGHT: white 1px solid; BACKGROUND: #d6dff7; OVERFLOW: hidden; BORDER-LEFT: white 1px solid; BORDER-BOTTOM: white 1px solid
}
.menu_title SPAN {
	FONT-WEIGHT: bold; LEFT: 8px; COLOR: #215dc6; POSITION: relative; TOP: 2px
}

.menu_title2 SPAN {
	FONT-WEIGHT: bold; LEFT: 8px; COLOR: #428eff; POSITION: relative; TOP: 2px
}
</style>
<div style="height:100%;OVERFLOW:auto">
	<table cellspacing="0" cellpadding="0" width="145" align="center">
		<tr>
			<td valign="bottom" height="42">
			<img height="38" src="images/menu/title.gif" width="145"></td>
		</tr>
		<tr>
			<td class="menu_title" onmouseover="this.className='menu_title2';" onmouseout="this.className='menu_title';" background="images/menu/title_bg_quit.gif" height="25">
			<span><b><font color="215DC6"><%=Request.Cookies("username")%></font></b> 
			| <a target="_top" href="?menu=out"><font color="215DC6"><b>退出</a>
			</font></a></b></span></td>
		</tr>
	</table>
	<br>
	<table cellspacing="0" cellpadding="0" width="145" align="center">
		<tr>
			<td class="menu_title" id="menuTitle2" onmouseover="this.className='menu_title2';" onclick="showsubmenu(1)" onmouseout="this.className='menu_title';" style="cursor:hand" background="images/menu/admin_left_2.gif" height="25">
			<span>设置管理</span> </td>
		</tr>
		<tr>
			<td id="submenu1" style="DISPLAY: none">
			<div class="sec_menu" style="WIDTH: 145px">
				<table cellspacing="0" cellpadding="0" width="135" align="center">
					<tr>
						<td height="20">
						<a target="main" href="admin_setup.asp?menu=variable">
						<font color="000000">社区基本设置</font></a></td>
					</tr>
					<tr>
						<td height="19">
						<a target="main" href="admin_setup.asp?menu=affichelist">
						<font color="000000">发布社区公告</font></a></td>
					</tr>
				</table>
			</div>
			<br>
			</td>
		</tr>
		<tr>
			<td class="menu_title" id="menuTitle1" onmouseover="this.className='menu_title2';" onclick="showsubmenu(2)" onmouseout="this.className='menu_title';" style="cursor:hand" background="images/menu/admin_left_3.gif" height="25">
			<span>用户管理 </span></td>
		</tr>
		<tr>
			<td id="submenu2" style="DISPLAY: none">
			<div class="sec_menu" style="WIDTH: 145px">
				<table cellspacing="0" cellpadding="0" width="135" align="center">
					<tr>
						<td height="20">
						<a target="main" href="admin_user.asp?menu=agreement">
						<font color="000000">注册用户协议</font></a></td>
					</tr>
					<tr>
						<td height="20">
						<a target="main" href="admin_user.asp?menu=showall">
						<font color="000000">注册用户管理</font></a></td>
					</tr>
					<tr>
					<td height="20">
					<a target="main" href="admin_user.asp?menu=user">
					<font color="000000">用户发放工资</font></a></td>
				</tr>
					<tr>
						<td height="20">
						<a target="main" href="admin_user.asp?menu=Activation">
						<font color="000000">激活用户资料</font></a></td>
					</tr>
				</table>
			</div>
			<br>
			</td>
		</tr>
	</table>
	<table cellspacing="0" cellpadding="0" width="145" align="center">
		<tr>
			<td class="menu_title" id="menuTitle1" onmouseover="this.className='menu_title2';" onclick="showsubmenu(3)" onmouseout="this.className='menu_title';" style="cursor:hand" background="images/menu/admin_left_4.gif" height="25">
			<span>论坛管理 </span></td>
		</tr>
		<tr>
			<td id="submenu3" style="DISPLAY: none">
			<div class="sec_menu" style="WIDTH: 145px">
				<table cellspacing="0" cellpadding="0" width="135" align="center">
					<tr>
						<td height="20">
						<a target="main" href="admin_bbs.asp?menu=classs">
						<font color="000000">建立论坛数据</font></a></td>
					</tr>
					<tr>
						<td height="20">
						<a target="main" href="admin_bbs.asp?menu=bbsmanage">
						<font color="000000">管理论坛资料</font></a></td>
					</tr>
					<tr>
						<td height="20">
						<a target="main" href="admin_bbs.asp?menu=applymanage">
						<font color="000000">管理所有论坛</font></a></td>
					</tr>
					<tr>
						<td height="20">
						<a target="main" href="admin_bbs.asp?menu=activation">
						<font color="000000">激活论坛帖子</font></a></td>
					</tr>
					<tr>
						<td height="20">
						<a target="main" href="admin_bbs.asp?menu=upclubconfig">
						<font color="000000">更新论坛资料</font></a></td>
					</tr>
				</table>
			</div>
			<br>
			</td>
		</tr>
	</table>
	<table cellspacing="0" cellpadding="0" width="145" align="center">
		<tr>
			<td class="menu_title" id="menuTitle1" onmouseover="this.className='menu_title2';" onclick="showsubmenu(4)" onmouseout="this.className='menu_title';" style="cursor:hand" background="images/menu/admin_left_5.gif" height="25">
			<span>社区管理</span> </td>
		</tr>
		<tr>
			<td id="submenu4" style="DISPLAY: none">
			<div class="sec_menu" style="WIDTH: 145px">
				<table cellspacing="0" cellpadding="0" width="135" align="center">
					<tr>
						<td height="19">
						<a target="main" href="admin_club.asp?menu=sendmail">
						<font color="000000">群发邮件</font></a></td>
					</tr>
					<tr>
						<td height="20">
						<a target="main" href="admin_club.asp?menu=message">
						<font color="000000">短讯息管理</font></a></td>
					</tr>
					<tr>
						<td height="20">
						<a target="main" href="admin_club.asp?menu=link">
						<font color="000000">友情链接管理</font></a></td>
					</tr>
					<tr>
						<td height="20">
						<a target="main" href="admin_club.asp?menu=faction">
						<font color="000000">社区帮派管理</font></a></td>
					</tr>
				</table>
			</div>
			<br>
			</td>
		</tr>
	</table>
	<table cellspacing="0" cellpadding="0" width="145" align="center">
		<tr>
			<td class="menu_title" id="menuTitle1" onmouseover="this.className='menu_title2';" onclick="showsubmenu(5)" onmouseout="this.className='menu_title';" style="cursor:hand" background="images/menu/admin_left_8.gif" height="25">
			<span>菜单管理</span></td>
		</tr>
		<tr>
			<td id="submenu5" style="DISPLAY: none">
			<div class="sec_menu" style="WIDTH: 145px">
				<table cellspacing="0" cellpadding="0" width="135" align="center">
					<tr>
						<td height="20"><a target="main" href="admin_menu.asp">
						<font color="000000">论坛菜单管理</font></a></td>
					</tr>
				</table>
			</div>
			<br>
			</td>
		</tr>
	</table>
	<table cellspacing="0" cellpadding="0" width="145" align="center">
		<tr>
			<td class="menu_title" id="menuTitle1" onmouseover="this.className='menu_title2';" onclick="showsubmenu(6)" onmouseout="this.className='menu_title';" style="cursor:hand" background="images/menu/admin_left_1.gif" height="25">
			<span>广告管理</span></td>
		</tr>
		<tr>
			<td id="submenu6" style="DISPLAY: none">
			<div class="sec_menu" style="WIDTH: 145px">
				<table cellspacing="0" cellpadding="0" width="135" align="center">
					<tr>
						<td height="20">
						<a target="main" href="admin_ad.asp?menu=banner">
						<font color="000000">社区广告设置</font></a></td>
					</tr>
				</table>
			</div>
			<br>
			</td>
		</tr>
	</table>
	<table cellspacing="0" cellpadding="0" width="145" align="center">
		<tr>
			<td class="menu_title" id="menuTitle1" onmouseover="this.className='menu_title2';" onclick="showsubmenu(7)" onmouseout="this.className='menu_title';" style="cursor:hand" background="images/menu/admin_left_7.gif" height="25">
			<span>数据管理 </span></td>
		</tr>
		<tr>
			<td id="submenu7" style="DISPLAY: none">
			<div class="sec_menu" style="WIDTH: 145px">
				<table cellspacing="0" cellpadding="0" width="135" align="center">
					<tr>
						<td height="20">
						<a target="main" href="admin_fso.asp?menu=statroom">
						<font color="000000">统计占用空间</font></a></td>
					</tr>
					<tr>
						<td height="20">
						<a target="main" href="admin_fso.asp?menu=files">
						<font color="000000">上传附件管理</font></a></td>
					</tr>
					<tr>
						<td height="20">
						<a target="main" href="admin_fso.asp?menu=ReForum">
						<font color="000000">帖子数据管理</font></a> </td>
					</tr>
					<tr>
						<td height="20">
						<a target="main" href="admin_fso.asp?menu=bak">
						<font color="000000">ACCESS数据库</font></a> </td>
					</tr>
				</table>
			</div>
			<br>
			</td>
		</tr>
	</table>
	<table cellspacing="0" cellpadding="0" width="145" align="center">
		<tr>
			<td class="menu_title" id="menuTitle1" onmouseover="this.className='menu_title2';" onclick="showsubmenu(8)" onmouseout="this.className='menu_title';" style="cursor:hand" background="images/menu/admin_left_6.gif" height="25">
			<span>其他功能</span> </td>
		</tr>
		<tr>
			<td id="submenu8" style="DISPLAY: none">
			<div class="sec_menu" style="WIDTH: 145px">
				<table cellspacing="0" cellpadding="0" width="135" align="center">
					<tr>
						<td height="20">
						<a target="main" href="admin_other.asp?menu=circumstance">
						<font color="000000">主机环境变量</font></a></td>
					</tr>
					<tr>
						<td height="20">
						<a target="main" href="admin_other.asp?menu=discreteness">
						<font color="000000">组件支持情况</font></a></td>
					</tr>
					<tr>
						<td height="20">
						<a target="main" href="admin_other.asp?menu=log">
						<font color="000000">系统操作记录</font></a></td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
	</table>
	<br>
</div>
<script>
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
}
}
</script>
<%
end sub
sub index


%>
<style type="text/css">.navPoint {COLOR: white; CURSOR: hand; FONT-FAMILY: Webdings; FONT-SIZE: 9pt}</style>
<script>
function switchSysBar(){
if (switchPoint.innerText==3){
switchPoint.innerText=4
document.all("frmTitle").style.display="none"
}else{
switchPoint.innerText=3
document.all("frmTitle").style.display=""
}}
</script>
<body scroll="no" topmargin="0" rightmargin="0" leftmargin="0">

<table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
	<tr>
		<td align="middle" id="frmTitle" nowrap valign="center" name="frmTitle">
		<iframe frameborder="0" id="carnoc" name="carnoc" scrolling="no" src="?menu=left" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 170px; Z-INDEX: 2">
		</iframe></td>
		<td class="a2" style="WIDTH: 9pt">
		<table border="0" cellpadding="0" cellspacing="0" height="100%">
			<tr>
				<td style="HEIGHT: 100%" onclick="switchSysBar()">
				<font style="FONT-SIZE: 9pt; CURSOR: default; COLOR: #ffffff">
				<br>
				<br>
				<br>
				<br>
				<br>
				<br>
				<br>
				<br>
				<br>
				<br>
				<br>
				<br>
				<br>
				<span class="navPoint" id="switchPoint" title="关闭/打开左栏">3</span><br>
				<br>
				<br>
				<br>
				<br>
				<br>
				<br>
				<br>
				屏幕切换 </font></td>
			</tr>
		</table>
		</td>
		<td style="WIDTH: 100%">
		<iframe frameborder="0" id="main" name="top" scrolling="no" scrolling="yes" src="?menu=top" style="HEIGHT: 4%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1">
		</iframe>
		<iframe frameborder="0" id="main" name="main" scrolling="yes" src="?menu=login" style="HEIGHT: 96%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1">
		</iframe></td>
	</tr>
</table>
</center>

</html>
<%
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

%>