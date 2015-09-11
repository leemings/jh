<!-- #include file="setup.asp" -->
<style type="text/css">
BODY {
	BACKGROUND:799ae1; MARGIN: 0px;
}

.sec_menu {
	BORDER-RIGHT: white 1px solid; BACKGROUND: #d6dff7; OVERFLOW: hidden; BORDER-LEFT: white 1px solid; BORDER-BOTTOM: white 1px solid
}

.menu_title SPAN {
	FONT-WEIGHT: bold; LEFT: 10px; COLOR: #215dc6; POSITION: relative; TOP: 2px 
}

.menu_title2 SPAN {
	FONT-WEIGHT: bold; LEFT: 10px; COLOR: #428eff; POSITION: relative; TOP: 2px
}
</style>
<table cellspacing="0" cellpadding="0" width="158" align="center">
	<tr>
		<td valign="bottom" height="42">
		<img height="38" src="images/title.gif" width="158" border="0"></td>
	</tr>
	<tr>
		<td class="menu_title" onmouseover="this.className='menu_title2';" onmouseout="this.className='menu_title';" background="images/title_bg_quit.gif" height="25">
		<span><a target="main" href="Default.asp"><b><font color="215DC6">首页</font></b></a> 
		|
		<script>
var undefined
if (getCookie('username')==undefined){
document.write("<a target=main href=login.asp><font color=215DC6><B>登录</a>")
}else
{
document.write("<a target=main href=login.asp?menu=out><font color=215DC6><B>退出</a>")
}
</script>
		</font></a></b></span></td>
	</tr>
	<tr>
		<td align="center" onmouseover="aa('up')" onmouseout="StopScroll()">
		<font face="Webdings" color="#FFFFFF" style=cursor:hand>5</font> </td>
	</tr>
</table>
<script>
var he=document.body.clientHeight-105
document.write("<div id=tt style=height:"+he+";overflow:hidden>")
</script>
<table cellspacing="0" cellpadding="0" width="158" align="center">
	<tr>
		<td id="imgmenu1" class="menu_title" onmouseover="this.className='menu_title2';" onclick="showsubmenu(1);loadThreadFollow()" onmouseout="this.className='menu_title';" style="cursor:hand" background=images/menudown.gif height="25">
		<span>论坛列表 </span></td>
	</tr>
	<tr>
		<td id="submenu1" style="DISPLAY: none">
		<div class="sec_menu">
			<table cellspacing="0" cellpadding="5" width="100%" style="border-collapse: collapse">
				<tr>
					<td height="20" id="followTd">Loading...</td>
				</tr>
			</table>
		</div>
		<br>
		</td>
	</tr>
	<tr>
		<td id="imgmenu2" class="menu_title" onmouseover="this.className='menu_title2';" onclick="showsubmenu(2)" onmouseout="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>控制面板 </span></td>
	</tr>
	<tr>
		<td id="submenu2" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr>
					<td><a target="main" href="usercp.asp">控制面板首页</a></td>
				</tr>
				<tr>
					<td><a target="main" href="EditProfile.asp">基本资料修改</a></td>
				</tr>
				<tr>
					<td><a target="main" href="EditProfile.asp?menu=pass">用户密码修改</a></td>
				</tr>
				<tr>
					<td><a target="main" href="EditProfile.asp?menu=option">编辑论坛选项</a></td>
				</tr>
				<tr>
					<td><a target="main" href="message.asp">用户短信服务</a></td>
				</tr>
				<tr>
					<td><a target="main" href="friend.asp">编辑好友列表</a></td>
				</tr>
			</table>
		</div>
		<br>
		</td>
	</tr>
	<tr>
		<td class="menu_title" id="imgmenu3" onmouseover="this.className='menu_title2';" onclick="showsubmenu(3)" onmouseout="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>个人服务</span> </td>
	</tr>
	<tr>
		<td id="submenu3" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<div align="center">
				<table cellspacing="3" cellpadding="0" width="130">
					<tr>
						<td><a target="main" href="login.asp">更换用户</a></td>
						<td><a target="main" href="ShowBBS.asp?menu=5">我的帖子</a></td>
					</tr>
					<tr>
						<td><a target="main" href="calendar.asp">我的日记</a></td>
						<td><a target="main" href="profile.asp">我的资料</a></td>
					</tr>
				</table>
			</div>
		</div>
		<br>
		</td>
	</tr>
	<tr>
		<td class="menu_title" id="imgmenu4" onmouseover="this.className='menu_title2';" onclick="showsubmenu(4)" onmouseout="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>个人收藏</span> </td>
	</tr>
	<tr>
		<td id="submenu4" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<div align="center">
				<table cellspacing="3" cellpadding="0" width="130">
					<tr>
						<td><a target="main" href="favorites.asp?menu=topic">帖子收藏夹</a></td>
					</tr>
					<tr>
						<td><a target="main" href="favorites.asp?menu=forum">论坛收藏夹</a></td>
					</tr>
					<tr>
						<td><a target="main" href="favorites.asp">网站收藏夹</a></td>
					</tr>
				</table>
			</div>
		</div>
		<br>
		</td>
	</tr>


	<%
sql="Select * From menu where followid=0 order by SortNum"
Set Rs=Conn.Execute(sql)
do while not rs.eof
%>
	<tr>
		<td class="menu_title" id="imgmenu<%=rs("id")%>0" onmouseover="this.className='menu_title2';" onclick="showsubmenu(<%=rs("id")%>0);loadingmenu(<%=rs("id")%>)" onmouseout="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span><%=rs("name")%></span> </td>
	</tr>
	<tr>
		<td id="submenu<%=rs("id")%>0" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px" align="center">
			<span id="menu<%=rs("id")%>">Loading...</span></div>
		<br>
		</td>
	</tr>

<%
rs.movenext
loop
Set Rs = Nothing
%>

	
	
	<tr>
		<td class="menu_title" id="imgmenu5" onmouseover="this.className='menu_title2';" onclick="showsubmenu(5)" onmouseout="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>社区概括</span> </td>
	</tr>
	<tr>
		<td id="submenu5" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<div align="center">
				<table cellspacing="3" cellpadding="0" width="130">
					<tr>
						<td></span><a target="main" href="online.asp">在线情况</a></td>
						<td><a target="main" href="online.asp?menu=cutline">在线图例</a></td>
					</tr>
					<tr>
						<td><a target="main" href="online.asp?menu=sex">性别图例</a></td>
						<td><a target="main" href="online.asp?menu=todaypage">今日图例</a></td>
					</tr>
					<tr>
						<td><a target="main" href="online.asp?menu=board">主题图例</a></td>
						<td><a target="main" href="online.asp?menu=tolrestore">帖子图例</a></td>
					</tr>
					<tr>
						<td><a target="main" href="usertop.asp">会员列表</a></td>
						<td><a target="main" href="adminlist.asp">管理团队</a></td>
					</tr>
				</table>
			</div>
		</div>
		</td>
	</tr>
</table>
&nbsp;
<table cellspacing="0" cellpadding="0" width="158" align="center">
	<tr>
		<td class="menu_title" onmouseover="this.className='menu_title2';" onmouseout="this.className='menu_title';" background="images/title_bg_quit.gif" height="25">
		<span>BBSxp Information</span> </td>
	</tr>
	<tr>
		<td>
		<div class="sec_menu" style="WIDTH: 158px">
			<div align="center">
			<table cellspacing="4" cellpadding="3">
				<tr>
					<td height="20">Version:<a target=_blank href=http://www.bbsxp.com/download.htm><font color=000000><%=ver%></font></a><br>
					Copyright:<a target="_blank" href="http://www.yuzi.net/"><font color="000000">Yuzi Studio</font></a><br>
					</td>
				</tr>
			</table>
			</div>
		</div>
		</td>
	</tr>
</table>
</div>
<table cellspacing="0" cellpadding="0" width="158" align="center">
	<tr>
		<td align="center" onmouseover="aa('Down')" onmouseout="StopScroll()" valign="bottom">
		<font face="Webdings" color="#FFFFFF" style=cursor:hand>6</font> </td>
	</tr>
</table>
<script>

function aa(Dir)
{tt.doScroll(Dir);Timer=setTimeout('aa("'+Dir+'")',100)}//这里100为滚动速度
function StopScroll(){if(Timer!=null)clearTimeout(Timer)}



function initIt(){
divColl=document.all.tags("DIV");
for(i=0; i<divColl.length; i++) {
whichEl=divColl(i);
if(whichEl.className=="child")whichEl.style.display="none";}
}
function expands(el) {
whichEl1=eval(el+"Child");
if (whichEl1.style.display=="none"){
initIt();
whichEl1.style.display="block";
}else{whichEl1.style.display="none";}
}
var tree= 0;
function loadThreadFollow(){
if (tree==0){
document.frames["hiddenframe"].location.replace("LeftTree.asp");
tree=1
}
}







function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
imgmenu = eval("imgmenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
imgmenu.background="images/menuup.gif";
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
imgmenu.background="images/menudown.gif";
}
}



function loadingmenu(id){
var loadmenu =eval("menu" + id);
if (loadmenu.innerText=="Loading..."){
document.frames["hiddenframe"].location.replace("LeftTree.asp?menu=menu&id="+id+"");
}
}
top.document.title="<%=clubname%> - Powered By BBSxp"; 
</script>
<%responseend%>