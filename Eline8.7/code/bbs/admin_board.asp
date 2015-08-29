<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file=md5.asp-->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	dim str
	dim admin_flag
	admin_flag="11,12"
	if not master or instr(session("flag"),"11")=0 or instr(session("flag"),"12")=0 then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		call main()
		conn.close
		set conn=nothing
	end if

	sub main()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" class="tableHeaderText" colspan=2 height=25>论坛管理
</th>
</tr>
<tr>
<td class="forumrow" colspan=2>
<p><B>注意</B>：<BR>①删除论坛同时将删除该论坛下所有帖子！删除分类同时删除下属论坛和其中帖子！ 操作时请完整填写表单信息。<BR>②如果选择<B>复位所有版面</B>，则所有版面都将作为一级论坛（分类），这时您需要重新对各个版面进行归属的基本设置，<B>不要轻易使用该功能</B>，仅在做出了错误的设置而无法复原版面之间的关系和排序的时候使用
</td>
</tr>
<tr>
<td class="forumrow">
<B>论坛操作选项</B></td>
<td class="forumrow"><a href="admin_board.asp">论坛管理首页</a> | <a href="admin_board.asp?action=add">新建论坛版面</a> | <a href="?action=orders">一级分类排序</a> | <a href="?action=boardorders">N级分类排序</a> | <a href="?action=RestoreBoard" onclick="{if(confirm('复位所有版面将把所有版面恢复成为一级大分类，复位后要对所有版面重新进行归属的基本设置，请慎重操作，确定复位吗?')){return true;}return false;}">复位所有版面</a>
</td>
</tr>
</table>
<p></p>
<%
select case Request("action")
case "add"
	call add()
case "edit"
	call edit()
case "savenew"
	call savenew()
case "savedit"
	call savedit()
case "del"
	call del()
case "orders"
	call orders()
case "updatorders"
	call updateorders()
case "boardorders"
	call boardorders()
case "updatboardorders"
	call updateboardorders()
case "addclass"
	call addclass()
case "saveclass"
	call saveclass()
case "del1"
	call del1()
case "mode"
	call mode()
case "savemod"
	call savemod()
case "permission"
	call boardpermission()
case "editpermission"
	call editpermission()
case "RestoreBoard"
	call RestoreBoard()
case else
	call boardinfo()
end select
end sub

sub boardinfo()
Dim reBoard_Setting
%>
<table width="95%" cellspacing="1" cellpadding="1" align=center class="tableBorder">
<tr> 
<th width="35%" class="tableHeaderText" height=25>论坛版面
</th>
<th width="35%" class="tableHeaderText" height=25>操作
</th>
</tr>
<%
sql="select * from board order by rootid,orders"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
do while not rs.eof
reBoard_Setting=split(rs("Board_setting"),",")
%>
<tr> 
<td height="25" width=35%  class="forumrow">
<%if rs("depth")>0 then%>
<%for i=1 to rs("depth")%>
&nbsp;
<%next%>
<%end if%>
<%if rs("child")>0 then%><img src="pic/plus.gif"><%else%><img src="pic/nofollow.gif"><%end if%>
<%if rs("parentid")=0 then%><b><%end if%><%=rs("boardtype")%><%if rs("child")>0 then%>(<%=rs("child")%>)<%end if%>
</td>
<td width=65% align=right class="forumrow"><a href="admin_board.asp?action=add&editid=<%=rs("boardid")%>"><font color="<%=Forum_body(14)%>"><U>添加版面</U></font></a> | <a href="admin_board.asp?action=edit&editid=<%=rs("boardid")%>"><font color="<%=Forum_body(14)%>"><U>基本设置</U></font></a> | <a href="admin_BoardSetting.asp?editid=<%=rs("boardid")%>"><font color="<%=Forum_body(14)%>"><U>高级设置</U></font></a>
<%if reBoard_Setting(2)=1 then%>
| <a href="admin_board.asp?action=mode&boardid=<%=rs("boardid")%>"><font color="<%=Forum_body(14)%>"><U>认证用户</U></font></a>
<%end if%>
| <a href="admin_update.asp?action=updat&submit=更新论坛数据&boardid=<%=rs("boardid")%>" title="更新最后回复、帖子数、回复数"><font color="<%=Forum_body(14)%>"><U>更新数据</U></font></a> | <a href="admin_update.asp?action=delboard&boardid=<%=rs("boardid")%>" onclick="{if(confirm('清空将包括该论坛所有帖子置于回收站，确定清空吗?')){return true;}return false;}"><font color="<%=Forum_body(14)%>"><U>清空</U></font></a> | <%if rs("child")=0 then%><a href="admin_board.asp?action=del&editid=<%=rs("boardid")%>" onclick="{if(confirm('删除将包括该论坛的所有帖子，确定删除吗?')){return true;}return false;}"><font color="<%=Forum_body(14)%>"><U>删除</U></a><%else%><a href="#" onclick="{if(confirm('该论坛含有下属论坛，必须先删除其下属论坛方能删除本论坛！')){return true;}return false;}"><font color="<%=Forum_body(14)%>"><U>删除</U></a><%end if%>&nbsp;</td>
</tr>
<%
rs.movenext
loop
set rs=nothing
%>
</table><BR><BR>
<%
end sub

sub add()
dim rs_c
set rs_c= server.CreateObject ("adodb.recordset")
sql = "select * from board order by rootid,orders"
rs_c.open sql,conn,1,1
	dim boardnum
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select Max(boardid) from board"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
	boardnum=1
	else
	boardnum=rs(0)+1
	end if
	if isnull(boardnum) then boardnum=1
	rs.close
%>
<form action ="admin_board.asp?action=savenew" method=post>
<input type="hidden" name="newboardid" value=<%=boardnum%>>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=24 colspan=2><B>添加新论坛</th>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow">论坛名称</td>
<td width="48%" class="forumrow"> 
<input type="text" name="boardtype" size="35">
</td>
</tr>
<tr> 
<td width="52%" height=24 class="forumrow">版面说明<BR>可以使用HTML代码</td>
<td width="48%" class="forumrow"> 
<textarea name="Readme" cols="40" rows="3"></textarea>
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>所属类别</U></td>
<td width="48%" class="forumrow"> 
<select name=class>
<option value="0">做为论坛分类</option>
<% do while not rs_c.EOF%>
<option value="<%=rs_c("boardid")%>" <%if request("editid")<>"" and clng(request("editid"))=rs_c("boardid") then%>selected<%end if%>>
<%if rs_c("depth")>0 then%>
<%for i=1 to rs_c("depth")%>
－
<%next%>
<%end if%><%=rs_c("boardtype")%></option>
<%
rs_c.MoveNext 
loop
rs_c.Close 
%>
</select>
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>使用设置模板</U><BR>相关模板中包含论坛颜色、设置、广告、图片<BR>等信息</td>
<td width="48%" class="forumrow"> 
<select name=sid>
<%
	sql = "select * from config order by active desc"
	rs_c.open sql,conn,1,1
	if rs_c.eof and rs_c.bof then
	response.write "<option value=>请先添加模板"
	else
	do while not rs_c.EOF
%>
<option value="<%=rs_c("id")%>"><%=rs_c("skinname")%> 
<%
	rs_c.MoveNext 
	loop
	end if
	rs_c.Close 
%>
</select>
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>论坛版主</U><BR>多版主添加请用|分隔，如：沙滩小子|wodeail</td>
<td width="48%" class="forumrow"> 
<input type="text" name="boardmaster" size="35">
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>首页显示论坛图片</U><BR>出现在首页论坛版面介绍左边<BR>请直接填写图片URL</td>
<td width="48%" class="forumrow">
<input type="text" name="indexIMG" size="35">
</td>
</tr>
<tr> 
<td width="52%" height=24 class="forumRow">&nbsp;</td>
<td width="48%" class="forumRow"> 
<input type="submit" name="Submit" value="添加论坛">
</td>
</tr>
</table>
</form>
<%
set rs_c=nothing
set rs=nothing
end sub

sub edit()
dim rs_c
sql = "select * from board order by rootid,orders"
set rs_c=conn.execute(sql)
sql = "select * from board where boardid="&request("editid")
set rs=conn.execute(sql)
%>        
 <form action ="admin_board.asp?action=savedit" method=post>       
<input type="hidden" name=editid value="<%=Request("editid")%>">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=24 colspan=2>编辑论坛：<%=rs("boardtype")%></th>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow">论坛名称</td>
<td width="48%" class="forumrow"> 
<input type="text" name="boardtype" size="35"  value="<%=rs("boardtype")%>">
</td>
</tr>
<tr> 
<td width="52%" height=24 class="forumrow">版面说明<BR>可以使用HTML代码</td>
<td width="48%" class="forumrow"> 
<textarea name="Readme" cols="40" rows="3"><%=rs("readme")%></textarea>
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>所属类别</U><BR>所属论坛不能指定为当前版面<BR>所属论坛不能指定为当前版面的下属论坛</td>
<td width="48%" class="forumrow"> 
<select name=class>
<option value="0">做为论坛分类</option>
<% do while not rs_c.EOF%>
<option value="<%=rs_c("boardid")%>" <% if cint(rs("parentid")) = rs_c("boardid") then%> selected <%end if%>><%if rs_c("depth")>0 then%>
<%for i=1 to rs_c("depth")%>
－
<%next%>
<%end if%><%=rs_c("boardtype")%></option>
<%
rs_c.MoveNext 
loop
rs_c.Close 
%>
</select>
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>使用设置模板</U><BR>相关模板中包含论坛颜色、设置、广告、图片<BR>等信息</td>
<td width="48%" class="forumrow"> 
<select name=sid>
<%
	sql = "select * from config order by active desc"
	rs_c.open sql,conn,1,1
	if rs_c.eof and rs_c.bof then
	response.write "<option value=>请先添加模板"
	else
	do while not rs_c.EOF
%>
<option value=<%=rs_c("id")%> <% if cint(rs("sid")) = rs_c("id") then%> selected <%end if%>><%=rs_c("skinname")%> 
<%
	rs_c.MoveNext 
	loop
	end if
	rs_c.Close 
%>
</select>
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>论坛版主</U><BR>多版主添加请用|分隔，如：沙滩小子|wodeail</td>
<td width="48%" class="forumrow"> 
<input type="text" name="boardmaster" size="35"  value='<%=rs("boardmaster")%>'>
<input type="hidden" name="oldboardmaster" value='<%=rs("boardmaster")%>'>
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>首页显示论坛图片</U><BR>出现在首页论坛版面介绍左边<BR>请直接填写图片URL</td>
<td width="48%" class="forumrow">
<input type="text" name="indexIMG" size="35" value="<%=rs("indexIMG")%>">
</td>
</tr>
<tr> 
<td width="52%" height=24 class="forumrow">&nbsp;</td>
<td width="48%" class="forumrow"> 
<input type="submit" name="Submit" value="提交修改">
</td>
</tr>
</table>
</form>
<%
rs.close
set rs=nothing
set rs_c=nothing
end sub
sub mode()
dim boarduser
%>
 <form action ="admin_board.asp?action=savemod" method=post>
<table width="95%" class="tableBorder" cellspacing="1" cellpadding="1" align="center">
<tr bgcolor=<%=Forum_body(3)%>> 
<th width="52%" height=22>说明：</th>
<th width="48%">操作：</th>
</tr>
<tr> 
<td width="52%" height=22 class=forumrow><B>论坛名称</B></td>
<td width="48%" class=forumrow> 
<%
set rs= server.CreateObject ("adodb.recordset")
sql="select boardid,boardtype,boarduser from board where boardid="&request("boardid")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
response.write "该版面并不存在或者该版面不是加密版面。"
else
response.write rs(1)
response.write "<input type=hidden value="&rs(0)&" name=boardid>"
boarduser=rs(2)
end if
rs.close
set rs=nothing
%>
</td>
</tr>
<tr> 
<td width="52%" class=forumrow><B>认证用户</B>：<br>
只有设定为认证论坛的论坛需要填写能够进入该版面的用户，每输入一个用户请确认用户名在论坛中存在，每个用户名用<B>回车</B>分开</font>
</td>
<td width="48%" class=forumrow> 
<textarea cols=35 rows=6 name="vipuser">
<%
if not isnull(boarduser) or boarduser<>"" then
	response.write replace(boarduser,",",chr(10))
end if
%>
</textarea>
</td>
</tr>
<tr> 
<td width="52%" height=22 class=forumrow>&nbsp;</td>
<td width="48%" class=forumrow> 
<input type="submit" name="Submit" value="设 定">
</td>
</tr>
</table>
</form>
<%
end sub

'保存编辑论坛认证用户信息
'入口：用户列表字符串
sub savemod()
dim boarduser
dim boarduser_1
dim userlen
dim updateinfo
response.write "<p>论坛设置成功！<br><br>"
if trim(request("vipuser"))<>"" then
	boarduser=request("vipuser")
	boarduser=split(boarduser,chr(13)&chr(10))
	for i = 0 to ubound(boarduser)
	if not (boarduser(i)="" or boarduser(i)=" ") then
		boarduser_1=""&boarduser_1&""&boarduser(i)&","
	end if
	next
	userlen=len(boarduser_1)
	if boarduser_1<>"" then
		boarduser=left(boarduser_1,userlen-1)
		response.write "<p>添加用户："&boarduser&"<br><br>"
		updateinfo=" boarduser='"&boarduser&"' "
	else
		response.write "<p><font color=red>你没有添加认证用户</font><br><br>"
	end if
end if
conn.execute("update board set "&updateinfo&" where boardid="&request("boardid"))
end sub

'保存添加论坛信息
sub savenew()
if request("boardtype")="" then
	Errmsg=Errmsg+"<br>"+"<li>请输入论坛名称。"
	Founderr=true
end if
if request("class")="" then
	Errmsg=Errmsg+"<br>"+"<li>请选择论坛分类。"
	Founderr=true
end if
if request("readme")="" then
	Errmsg=Errmsg+"<br>"+"<li>请输入论坛说明。"
	Founderr=true
end if
if founderr=true then
	response.write Errmsg
	exit sub
end if
dim boardid
dim rootid
dim parentid
dim depth
dim orders
dim Fboardmaster
dim maxrootid
dim parentstr
if request("class")<>"0" then
set rs=conn.execute("select rootid,boardid,depth,orders,boardmaster,ParentStr from board where boardid="&request("class"))
rootid=rs(0)
parentid=rs(1)
depth=rs(2)
orders=rs(3)
if depth+1>20 then
	response.write "本论坛限制最多只能有20级分类"
	exit sub
end if
parentstr=rs(5)
else
set rs=conn.execute("select max(rootid) from board")
maxrootid=rs(0)+1
if isnull(MaxRootID) then MaxRootID=1
end if
sql="select boardid from board where boardid="&request("newboardid")
set rs=conn.execute(sql)
if not (rs.eof and rs.bof) then
	response.write "您不能指定和别的论坛一样的序号。"
	exit sub
else
	boardid=request("newboardid")
end if

set rs = server.CreateObject ("adodb.recordset")
sql = "select * from board"
rs.Open sql,conn,1,3
rs.AddNew
if request("class")<>"0" then
rs("depth")=depth+1
rs("rootid")=rootid
rs("orders") = Request.form("newboardid")
rs("parentid") = Request.Form("class")
if ParentStr="0" then
rs("ParentStr")=Request.Form("class")
else
rs("ParentStr")=ParentStr & "," & Request.Form("class")
end if
else
rs("depth")=0
rs("rootid")=maxrootid
rs("orders")=0
rs("parentid")=0
rs("child")=0
rs("parentstr")=0
end if
rs("boardid") = Request.form("newboardid")
rs("boardtype") = Request.Form("boardtype")
rs("readme") = Request.form("readme")
rs("lasttopicnum") = 0
rs("lastbbsnum") = 0
rs("lasttopicnum") = 0
rs("todaynum") = 0
rs("LastPost")="$0$"&Now()&"$$$$$"
rs("Board_Setting")="0,0,0,0,1,0,1,1,1,1,1,1,1,1,1,1,16240,3,300,gif|jpg|jpeg|bmp|png|rar|txt|zip|mid,0,0,0|24,1,0,300,20,10,9,12,1,10,10,0,0,0,0,1,5,0,1,4,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,1,1,1,1,1"
rs("sid")=request.form("sid")
if Request("boardmaster")<>"" then
	rs("boardmaster") = Request.form("boardmaster")
end if
if request.form("indexIMG")<>"" then
	rs("indexIMG")=request.form("indexIMG")
end if
rs.Update 
rs.Close
if Request("boardmaster")<>"" then call addmaster(Request("boardmaster"),"none",0)
if request("class")<>"0" then
if depth>0 then
	'当上级分类深度大于0的时候要更新其父类（或父类的父类）的版面数和相关排序
	for i=1 to depth
		'更新其父类版面数
		if parentid<>"" then
		conn.execute("update board set child=child+1 where boardid="&parentid)
		end if
		'得到其父类的父类的版面ID
		set rs=conn.execute("select parentid from board where boardid="&parentid)
		if not (rs.eof and rs.bof) then
			parentid=rs(0)
		end if
		'当循环次数大于1并且运行到最后一次循环的时候直接进行更新
		if i=depth and parentid<>"" then
		conn.execute("update board set child=child+1 where boardid="&parentid)
		end if
	next
	'更新该版面排序以及大于本需要和同在本分类下的版面排序序号
	conn.execute("update board set orders=orders+1 where rootid="&rootid&" and orders>"&orders)
	conn.execute("update board set orders="&orders&"+1 where boardid="&Request.form("newboardid"))
else
	'当上级分类深度为0的时候只要更新上级分类版面数和该版面排序序号即可
	conn.execute("update board set child=child+1 where boardid="&request("class"))
	set rs=conn.execute("select max(orders) from board where boardid="&Request.form("newboardid"))
	conn.execute("update board set orders="&rs(0)&"+1 where boardid="&Request.form("newboardid"))
end if
end if
response.write "<p>论坛添加成功！<br><B>该论坛目前高级设置为默认选项，建议您返回论坛管理中心重新设置该论坛的高级选项</B><BR><br>"&str
set rs=nothing
call cache_board()
end sub

'保存编辑论坛信息
sub savedit()
if clng(request("editid"))=clng(request("class")) then
	response.write "所属论坛不能指定自己"
	exit sub
end if
dim newboardid,maxrootid
dim parentid,boardmaster,depth,child,ParentStr,rootid,iparentid,iParentStr
dim trs,brs,mrs
set rs = server.CreateObject ("adodb.recordset")
sql = "select * from board where boardid="&request("editid")
rs.Open sql,conn,1,3
newboardid=rs("boardid")
parentid=rs("parentid")
iparentid=rs("parentid")
boardmaster=rs("boardmaster")
ParentStr=rs("ParentStr")
depth=rs("depth")
child=rs("child")
rootid=rs("rootid")
'判断所指定的论坛是否其下属论坛
if ParentID=0 then
	if clng(request("class"))<>0 then
	set trs=conn.execute("select rootid from board where boardid="&request("class"))
	if rootid=trs(0) then
		response.write "您不能指定该版面的下属论坛作为所属论坛"
		exit sub
	end if
	end if
else
	set trs=conn.execute("select boardid from board where (ParentStr='"&ParentStr&"' or ParentStr like '"&ParentStr&",%') and boardid="&request("class"))
	if not (trs.eof and trs.bof) then
		response.write "您不能指定该版面的下属论坛作为所属论坛"
		response.end
	end if
end if
if parentid=0 then
	parentid=rs("boardid")
	iparentid=0
end if
rs("boardtype") = Request.Form("boardtype")
'rs("parentid") = Request.Form("class")
rs("boardmaster") = Request("boardmaster")
rs("readme") = Request("readme")
rs("indexIMG")=request.form("indexIMG")
rs("sid")=request.form("sid")
rs.Update 
rs.Close
set rs=nothing
if request("oldboardmaster")<>Request("boardmaster") then call addmaster(Request("boardmaster"),request("oldboardmaster"),1)

set mrs=conn.execute("select max(rootid) from board")
Maxrootid=mrs(0)+1
'假如更改了所属论坛
'需要更新其原来所属版面信息，包括深度、父级ID、版面数、排序、继承版主等数据
'需要更新当前所属版面信息
'继承版主数据需要另写函数进行更新--取消，在前台可用boardid in parentstr来获得
dim k,nParentStr,mParentStr
dim ParentSql,boardcount
if clng(parentid)<>clng(request("class")) and not (iparentid=0 and cint(request("class"))=0) then
	'如果原来不是一级分类改成一级分类
	if iparentid>0 and cint(request("class"))=0 then
		'更新当前版面数据
		conn.execute("update board set depth=0,orders=0,rootid="&maxrootid&",parentid=0,parentstr='0' where boardid="&newboardid)
		ParentStr=ParentStr & ","
		set rs=conn.execute("select count(*) from board where ParentStr='"&ParentStr&"' or ParentStr like '"&ParentStr&",%'")
		boardcount=rs(0)
		if isnull(boardcount) then
		boardcount=1
		else
		boardcount=boardcount+1
		end if
		'更新其原来所属论坛版面数
		conn.execute("update board set child=child-"&boardcount&" where boardid="&iparentid)
		'更新其原来所属论坛数据，排序相当于剪枝而不需考虑
		for i=1 to depth
			'得到其父类的父类的版面ID
			set rs=conn.execute("select parentid from board where boardid="&iparentid)
			if not (rs.eof and rs.bof) then
				iparentid=rs(0)
				conn.execute("update board set child=child-"&boardcount&" where boardid="&iparentid)
			end if
		next
		if child>0 then
		'更新其下属论坛数据
		'有下属论坛，排序不需考虑，更新下属论坛深度和一级排序ID(rootid)数据
		'更新当前版面数据
		'ParentStr=ParentStr & ","		
		i=0
		set rs=conn.execute("select * from board where ParentStr='"&ParentStr&"' or ParentStr like '"&ParentStr&",%'")
		do while not rs.eof
		i=i+1
		mParentStr=replace(rs("ParentStr"),ParentStr,"")
		conn.execute("update board set depth=depth-"&depth&",rootid="&maxrootid&",ParentStr='"&mParentStr&"' where boardid="&rs("boardid"))
		rs.movenext
		loop
		end if
	elseif iparentid>0 and cint(request("class"))>0 then
	'将一个分论坛移动到其他分论坛下
	'获得所指定的论坛的相关信息
	set trs=conn.execute("select * from board where boardid="&request("class"))
	'得到其下属版面数
	ParentStr=ParentStr & ","
	set rs=conn.execute("select count(*) from board where ParentStr='"&ParentStr&"' or ParentStr like '"&ParentStr&",%'")
	boardcount=rs(0)
	if isnull(boardcount) then boardcount=1
	'在获得移动过来的版面数后更新排序在指定论坛之后的论坛排序数据
	conn.execute("update board set orders=orders + "&boardCount&" + 1 where rootid="&trs("rootid")&" and orders>"&trs("orders")&"")
	'更新当前版面数据
	conn.execute("update board set depth="&trs("depth")&"+1,orders="&trs("orders")&"+1,rootid="&trs("rootid")&",ParentID="&request("class")&",ParentStr='" & trs("parentstr") & "," & trs("boardid") & "' where boardid="&newboardid)
	i=1
	'如果有则更新下属版面数据
	'深度为原有深度加上当前所属论坛的深度
	set rs=conn.execute("select * from board where ParentStr='"&ParentStr&"' or ParentStr like '"&ParentStr&",%' order by orders")
	do while not rs.eof
	i=i+1
	iParentStr=trs("parentstr") & "," & trs("boardid") & "," & replace(rs("parentstr"),ParentStr,"")
	conn.execute("update board set depth=depth+"&trs("depth")&"-"&depth&"+1,orders="&trs("orders")&"+"&i&",rootid="&trs("rootid")&",ParentStr='"&iParentStr&"' where boardid="&rs("boardid"))
	rs.movenext
	loop
	ParentID=request("class")
	if rootid=trs("rootid") then
	'在同一分类下移动
	'更新所指向的上级论坛版面数，i为本次移动过来的版面数
	'更新其父类版面数
	conn.execute("update board set child=child+"&i&" where (not ParentID=0) and boardid="&parentid)
	for k=1 to trs("depth")
		'得到其父类的父类的版面ID
		set rs=conn.execute("select parentid from board where (not ParentID=0) and boardid="&parentid)
		if not (rs.eof and rs.bof) then
			parentid=rs(0)
			'更新其父类的父类版面数
			conn.execute("update board set child=child+"&i&" where (not ParentID=0) and  boardid="&parentid)
		end if
	next
	'更新其原父类版面数
	conn.execute("update board set child=child-"&i&" where (not ParentID=0) and boardid="&iparentid)
	'更新其原来所属论坛数据
	'response.write iparentid & "<br>"
	for k=1 to depth
		'得到其原父类的父类的版面ID
		set rs=conn.execute("select parentid from board where (not ParentID=0) and boardid="&iparentid)
		if not (rs.eof and rs.bof) then
			iparentid=rs(0)
			'response.write iparentid & "<br>"
			'更新其原父类的父类版面数
			conn.execute("update board set child=child-"&i&" where (not ParentID=0) and  boardid="&iparentid)
		end if
	next
	else
	'更新所指向的上级论坛版面数，i为本次移动过来的版面数
	'更新其父类版面数
	conn.execute("update board set child=child+"&i&" where boardid="&parentid)
	for k=1 to trs("depth")
		'得到其父类的父类的版面ID
		set rs=conn.execute("select parentid from board where boardid="&parentid)
		if not (rs.eof and rs.bof) then
			parentid=rs(0)
			'更新其父类的父类版面数
			conn.execute("update board set child=child+"&i&" where boardid="&parentid)
		end if
	next
	'更新其原父类版面数
	conn.execute("update board set child=child-"&i&" where boardid="&iparentid)
	'更新其原来所属论坛数据
	for k=1 to depth
		'得到其原父类的父类的版面ID
		set rs=conn.execute("select parentid from board where boardid="&iparentid)
		if not (rs.eof and rs.bof) then
			iparentid=rs(0)
			'更新其原父类的父类版面数
			conn.execute("update board set child=child-"&i&" where boardid="&iparentid)
		end if
	next
	end if 'end if rootid=trs("rootid") then
	else
	'如果原来是一级论坛改成其他论坛的下属论坛
	'得到所指定的论坛的相关信息
	set trs=conn.execute("select * from board where boardid="&request("class"))
	set rs=conn.execute("select count(*) from board where rootid="&rootid)
	boardcount=rs(0)
	'更新所指向的上级论坛版面数，i为本次移动过来的版面数
	ParentID=request("class")
	'更新其父类版面数
	conn.execute("update board set child=child+"&boardcount&" where boardid="&parentid)
	'response.write parentid & "-"&boardcount&"<br>"
	for k=1 to trs("depth")
		'得到其父类的父类的版面ID
		set rs=conn.execute("select parentid from board where boardid="&parentid)
		if not (rs.eof and rs.bof) then
			parentid=rs(0)
			'更新其父类的父类版面数
			conn.execute("update board set child=child+"&boardcount&" where boardid="&parentid)
		end if
	'response.write parentid & "-"&boardcount&"<br>"
	next
	'在获得移动过来的版面数后更新排序在指定论坛之后的论坛排序数据
	conn.execute("update board set orders=orders + "&boardCount&" + 1 where rootid="&trs("rootid")&" and orders>"&trs("orders")&"")
	i=0
	set rs=conn.execute("select * from board where rootid="&rootid&" order by orders")
	do while not rs.eof
	i=i+1
	if rs("parentid")=0 then
		if trs("ParentStr")="0" then
		parentstr=trs("boardid")
		else
		parentstr=trs("parentstr") & "," & trs("boardid")
		end if
	conn.execute("update board set depth=depth+"&trs("depth")&"+1,orders="&trs("orders")&"+"&i&",rootid="&trs("rootid")&",ParentStr='"&ParentStr&"',parentid="&request("class")&" where boardid="&rs("boardid"))
	else
		if trs("ParentStr")="0" then
		parentstr=trs("boardid") & "," & rs("parentstr")
		else
		parentstr=trs("parentstr") & "," & trs("boardid") & "," & rs("parentstr")
		end if
	conn.execute("update board set depth=depth+"&trs("depth")&"+1,orders="&trs("orders")&"+"&i&",rootid="&trs("rootid")&",ParentStr='"&ParentStr&"' where boardid="&rs("boardid"))
	end if
	rs.movenext
	loop
	end if
end if
response.write "<p>论坛修改成功！<br><br>"&str
set rs=nothing
set mrs=nothing
set trs=nothing
'cache版面数据
call cache_board()
'end cache
end sub

'删除版面，删除版面帖子，入口：版面ID
sub del()
'更新其上级版面论坛数，如果该论坛含有下级论坛则不允许删除
set rs=conn.execute("select ParentStr,child,depth from board where boardid="&Request("editid"))
if not (rs.eof and rs.bof) then
if rs(1)>0 then
	response.write "该论坛含有下属论坛，请删除其下属论坛后再进行删除本论坛的操作"
	exit sub
end if
'如果有上级版面，则更新数据
if rs(2)>0 then
	conn.execute("update board set child=child-1 where boardid in ("&rs(0)&")")
end if
sql = "delete from board where boardid="&Request("editid")
conn.execute(sql)
for i=0 to ubound(AllPostTable)
sql = "delete from "&AllPostTable(i)&" where boardid="&Request("editid")
conn.execute(sql)
next
end if
set rs=nothing
call cache_board()
response.write "<p>论坛删除成功！"
end sub

sub orders()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr> 
	<th height="22">论坛一级分类重新排序修改(请在相应论坛分类的排序表单内输入相应的排列序号)
	</th>
	</tr>

	<tr>
	<td class="Forumrow"><table width="50%">
<%
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select * from Board where ParentID=0 order by RootID"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "还没有相应的论坛分类。"
	else
		do while not rs.eof
		response.write "<form action=admin_board.asp?action=updatorders method=post><tr><td width=""50%"">"&rs("boardtype")&"</td>"
		response.write "<td width=""50%""><input type=text name=""OrderID"" size=4 value="""&rs("rootid")&"""><input type=hidden name=""cID"" value="""&rs("rootid")&""">&nbsp;&nbsp;<input type=submit name=Submit value=修改></td></tr></form>"
		rs.movenext
		loop
%>
</table>
<BR>&nbsp;<font color=red>请注意，这里一定<B>不能填写相同的序号</B>，否则非常难修复！</font>
<%
	end if
	rs.close
	set rs=nothing
%>
	</td>
	</tr>
</table>
<%
end sub

sub updateorders()
	dim cID,OrderID,ClassName
	'response.write request.form("cID")(1)
	'response.end
	cID=replace(request.form("cID"),"'","")
	OrderID=replace(request.form("OrderID"),"'","")
	set rs=conn.execute("select boardid from board where rootid="&orderid)
	if rs.eof and rs.bof then
	response.write "设置成功，请返回。"
	conn.execute("update board set rootid="&OrderID&" where rootid="&cID)
	else
	response.write "请不要和其他论坛设置相同的序号"
	end if
	call cache_board()
end sub


sub boardorders()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr> 
	<th height="22">论坛N级分类重新排序修改(请在相应论坛分类的排序表单内输入相应的排列序号)
	</th>
	</tr>
	<tr>
	<td class="Forumrow"><table width="90%">
<%
dim trs,uporders,doorders
set rs = server.CreateObject ("Adodb.recordset")
sql="select * from Board order by RootID,orders"
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	response.write "还没有相应的论坛分类。"
else
	do while not rs.eof
	response.write "<form action=admin_board.asp?action=updatboardorders method=post><tr><td width=""50%"">"
	if rs("depth")>0 then
	for i=1 to rs("depth")
		response.write "&nbsp;"
	next
	end if
	if rs("child")>0 then
		response.write "<img src=pic/plus.gif>"
	else
		response.write "<img src=pic/nofollow.gif>"
	end if
	if rs("parentid")=0 then
		response.write "<b>"
	end if
	response.write rs("boardtype")
	if rs("child")>0 then
		response.write "("&rs("child")&")"
	end if
	response.write "</td><td width=""50%"">"
	if rs("ParentID")>0 then
	'算出相同深度的版面数目，得到该版面在相同深度的版面中所处位置（之上或者之下的版面数）
	'所能提升最大幅度应为For i=1 to 该版之上的版面数
	set trs=conn.execute("select count(*) from board where ParentID="&rs("ParentID")&" and orders<"&rs("orders")&"")
	uporders=trs(0)
	if isnull(uporders) then uporders=0
	'所能降低最大幅度应为For i=1 to 该版之下的版面数
	set trs=conn.execute("select count(*) from board where ParentID="&rs("ParentID")&" and orders>"&rs("orders")&"")
	doorders=trs(0)
	if isnull(doorders) then doorders=0
	if uporders>0 then
		response.write "<select name=uporders size=1><option value=0>向上移动</option>"
		for i=1 to uporders
		response.write "<option value="&i&">"&i&"</option>"
		next
		response.write "</select>"
	end if
	if doorders>0 then
		if uporders>0 then response.write "&nbsp;"
		response.write "<select name=doorders size=1><option value=0>向下移动</option>"
		for i=1 to doorders
		response.write "<option value="&i&">"&i&"</option>"
		next
		response.write "</select>"
	end if
	if doorders>0 or uporders>0 then
	response.write "<input type=hidden name=""editID"" value="""&rs("boardid")&""">&nbsp;<input type=submit name=Submit value=修改>"
	end if
	end if
	response.write "</td></tr></form>"
	uporders=0
	doorders=0
	rs.movenext
	loop
	response.write "</table>"
end if
rs.close
set rs=nothing
%>
	</td>
	</tr>
</table>
<%
end sub

sub updateboardorders()
dim ParentID,orders,ParentStr,child
dim uporders,doorders,oldorders,trs,ii
if not isnumeric(request("editID")) then
	response.write "非法的参数！"
	exit sub
end if
if request("uporders")<>"" and not Cint(request("uporders"))=0 then
	if not isnumeric(request("uporders")) then
	response.write "非法的参数！"
	exit sub
	elseif Cint(request("uporders"))=0 then
	response.write "请选择要提升的数字！"
	exit sub
	end if
	'向上移动
	'要移动的论坛信息
	set rs=conn.execute("select ParentID,orders,ParentStr,child from board where boardid="&request("editID"))
	ParentID=rs(0)
	orders=rs(1)
	ParentStr=rs(2) & "," & request("editID")
	child=rs(3)
	i=0
	'response.write "select boardid,orders from board where ParentID="&ParentID&" and orders<"&orders&" order by orders desc<br>"
	if child>0 then
	set rs=conn.execute("select count(*) from board where ParentStr='"&ParentStr&"' or ParentStr like '"&ParentStr&",%'")
	oldorders=rs(0)
	else
	oldorders=0
	end if
	'和该论坛同级且排序在其之上的论坛－更新其排序，最末者为当前论坛排序号
	set rs=conn.execute("select boardid,orders,child,ParentStr from board where ParentID="&ParentID&" and orders<"&orders&" order by orders desc")
	do while not rs.eof
	i=i+1
	if Cint(request("uporders"))>=i then
		'response.write "update board set orders="&orders&" where boardid="&rs(0)&"<br>"
		if rs(2)>0 then
		ii=0
		set trs=conn.execute("select boardid,orders from board where ParentStr='"&rs(3)&","&rs(0)&"' or ParentStr like '"&rs(3)&","&rs(0)&",%' order by orders")
		if not (trs.eof and trs.bof) then
		do while not trs.eof
		ii=ii+1
		conn.execute("update board set orders="&orders&"+"&oldorders&"+"&ii&" where boardid="&trs(0))
		trs.movenext
		loop
		end if
		end if
		conn.execute("update board set orders="&orders&"+"&oldorders&" where boardid="&rs(0))
		if Cint(request("uporders"))=i then uporders=rs(1)
	end if
	orders=rs(1)
	rs.movenext
	loop
	'response.write "update board set orders="&uporders&" where boardid="&request("editID")
	'更新所要排序的论坛的序号
	conn.execute("update board set orders="&uporders&" where boardid="&request("editID"))
	'如果有下属论坛，则更新其下属论坛排序
	if child>0 then
	i=uporders
	set rs=conn.execute("select boardid from board where ParentStr='"&ParentStr&"' or ParentStr like '"&ParentStr&",%' order by orders")
	do while not rs.eof
	i=i+1
	conn.execute("update board set orders="&i&" where boardid="&rs(0))
	rs.movenext
	loop
	end if
	'response.end
	set rs=nothing
	set trs=nothing
elseif request("doorders")<>"" then
	if not isnumeric(request("doorders")) then
	response.write "非法的参数！"
	exit sub
	elseif Cint(request("doorders"))=0 then
	response.write "请选择要下降的数字！"
	exit sub
	end if
	set rs=conn.execute("select ParentID,orders,ParentStr,child from board where boardid="&request("editID"))
	ParentID=rs(0)
	orders=rs(1)
	ParentStr=rs(2) & "," & request("editID")
	child=rs(3)
	i=0
	if child>0 then
	set rs=conn.execute("select count(*) from board where ParentStr='"&ParentStr&"' or ParentStr like '"&ParentStr&",%'")
	oldorders=rs(0)
	else
	oldorders=0
	end if
	set rs=conn.execute("select boardid,orders,child,ParentStr from board where ParentID="&ParentID&" and orders>"&orders&" order by orders")
	do while not rs.eof
	i=i+1
	if Cint(request("doorders"))>=i then
		if rs(2)>0 then
		ii=0
		set trs=conn.execute("select boardid,orders from board where ParentStr='%"&rs(3)&","&rs(0)&"%' or ParentStr like '"&rs(3)&","&rs(0)&",%' order by orders")
		if not (trs.eof and trs.bof) then
		do while not trs.eof
		ii=ii+1
		'response.write "update board set orders="&orders&"+"&ii&" where boardid="&trs(0)&"－a<br>"
		conn.execute("update board set orders="&orders&"+"&ii&" where boardid="&trs(0))
		trs.movenext
		loop
		end if
		end if
		'response.write "update board set orders="&orders&" where boardid="&rs(0)&"<br>"
		conn.execute("update board set orders="&orders&" where boardid="&rs(0))
		if Cint(request("doorders"))=i then doorders=rs(1)
	end if
	orders=rs(1)
	rs.movenext
	loop
	'response.write "update board set orders="&doorders&" where boardid="&request("editID")&"<br>"
	conn.execute("update board set orders="&doorders&" where boardid="&request("editID"))
	'如果有下属论坛，则更新其下属论坛排序
	if child>0 then
	i=doorders
	set rs=conn.execute("select boardid from board where ParentStr='"&ParentStr&"' or ParentStr like '"&ParentStr&",%' order by orders")
	do while not rs.eof
	i=i+1
	'response.write "update board set orders="&i&" where boardid="&rs(0)&"－b<br>"
	conn.execute("update board set orders="&i&" where boardid="&rs(0))
	rs.movenext
	loop
	end if
	'response.end
	set rs=nothing
	set trs=nothing
end if
call cache_board()
response.redirect "admin_board.asp?action=boardorders"
end sub

sub addmaster(s,o,n)
dim arr,pw,oarr
dim classname,titlepic
set rs=conn.execute("select title from usergroups where usergroupid=3")
classname=rs(0)
set rs=conn.execute("select titlepic from usertitle where usergroupid=3 order by Minarticle desc")
if not (rs.eof and rs.bof) then
titlepic=rs(0)
end if
randomize
pw=Cint(rnd*9000)+1000
arr=split(s,"|")
oarr=split(o,"|")
set rs=server.createobject("adodb.recordset")
for i=0 to Ubound(arr)
sql="select * from [user] where username='"& arr(i) &"'"
rs.open sql,conn,1,3
if rs.eof and rs.bof then
	rs.addnew
	rs("username")=arr(i)
	rs("userpassword")=md5(pw)
	rs("userclass")=classname
	rs("UserGroupID")=3
	rs("titlepic")=titlepic
	rs("userWealth")=Forum_user(0)
	rs("userep")=Forum_user(5)
	rs("usercp")=Forum_user(10)
	rs("userisbest")=0
	rs("userdel")=0
	rs("userpower")=0
	rs("lockuser")=0
	rs.update
	str=str&"你添加了以下用户：<b>" &arr(i) &"</b> 密码：<b>"& pw &"</b><br><br>"
else
	if rs("UserGroupID")>3 then
		rs("userclass")=classname
		rs("UserGroupID")=3
		rs("titlepic")=titlepic
		rs.update
	end if
end if
rs.close
next

'判断原版主在其他版面是否还担任版主，如没有担任则撤换该用户职位
if n=1 then
dim iboardmaster
dim UserGrade,article
iboardmaster=false
for i=0 to ubound(oarr)
	set rs=conn.execute("select boardmaster from board")
	do while not rs.eof
		if instr("|"&lcase(trim(rs("boardmaster")))&"|","|"&lcase(trim(oarr(i)))&"|")>0 then
			iboardmaster=true
			exit do
		end if
	rs.movenext
	loop
	if not iboardmaster then
	set rs=conn.execute("select userid,UserGroupID,article from [user] where username='"&trim(oarr(i))&"'")
	if not (rs.eof and rs.bof) then
		if rs(1)>2 then
		if not isnumeric(rs(2)) then article=0
		set UserGrade=conn.execute("select top 1 usertitle,titlepic,UserGroupID from usertitle where Minarticle<="&rs(2)&" and not MinArticle=-1 order by MinArticle desc,usertitleid")
		if not (UserGrade.eof and UserGrade.bof) then
			conn.execute("update [user] set UserGroupID="&UserGrade(2)&",titlepic='"&UserGrade(1)&"',userclass='"&UserGrade(0)&"' where userid="&rs(0))
		end if
		end if
	end if
	end if
	iboardmaster=false
next
end if
set rs=nothing
end sub

sub boardpermission()
dim iUserGroupID(20),UserTitle(20)
dim trs,ars,k,ii
ii=0
set trs=conn.execute("select title,usergroupid from usergroups order by usergroupid")
do while not trs.eof
UserTitle(ii)=trs(0)
iUserGroupID(ii)=trs(1)
ii=ii+1
trs.movenext
loop
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr>
	<th height="25">编辑论坛权限</th>
	</tr>
	<tr>
	<td class=forumrow>①您可以设置不同用户组在不同论坛内的权限，红色表示为该论坛该用户组使用的是用户定义属性<BR>②该权限不能继承，比如您设置了一个包含下级论坛的版面，那么只对您设置的版面生效而不对其下属论坛生效<BR>③如果您想设置生效，必须在设置页面<B>选择自定义设置</B>，选择了自定义设置后，这里设置的权限将<B>优先</B>于用户组设置，比如用户组默认不能管理帖子，而这里设置了该用户组可管理帖子，那么该用户组在这个版面就可以管理帖子
	</td>
	</tr>
</table><BR>
<table width="95%" cellspacing="1" cellpadding="1" align=center class="tableBorder">
<tr> 
<th width="35%" class="tableHeaderText" height=25>论坛版面
</th>
<th width="35%" class="tableHeaderText" height=25>设置用户组权限
</th>
</tr>
<%
sql="select * from board order by rootid,orders"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
do while not rs.eof
%>
<tr> 
<td height="25" width=40%  class="forumrow">
<%if rs("depth")>0 then%>
<%for i=1 to rs("depth")%>
&nbsp;
<%next%>
<%end if%>
<%if rs("child")>0 then%><img src="pic/plus.gif"><%else%><img src="pic/nofollow.gif"><%end if%>
<%if rs("parentid")=0 then%><b><%end if%><%=rs("boardtype")%><%if rs("child")>0 then%>(<%=rs("child")%>)<%end if%>
</td>
<FORM METHOD=POST ACTION="?action=editpermission">
<td width=60% class="forumrow">&nbsp;
<select name="groupid" size=1>
<%
	for k=0 to ii-1
		set ars=conn.execute("select pid from BoardPermission where BoardID="&rs("boardid")&" and GroupID="&iUserGroupID(k))
		if ars.eof and ars.bof then
		response.write "<option value="""&iUserGroupID(k)&""">" & UserTitle(k) & "</option>"
		else
		response.write "<option value="""&iUserGroupID(k)&""">" & UserTitle(k) & "(自定义)</option>"
		end if
	next
%>
</select>
<input type=hidden value="<%=rs("boardid")%>" name=reboardid>
<input type=submit name=submit value="设置">
<%
	dim percount
	set trs=conn.execute("select count(*) from BoardPermission where boardid="&rs("boardid"))
	percount=trs(0)
	if not isnull(percount) and percount>0 then response.write "(有自定义版面)"

%>
</td>
</FORM>
</tr>
<%
rs.movenext
loop
set rs=nothing
%>
</table><BR><BR>
<%
set trs=nothing
set ars=nothing
end sub

sub editpermission()
if not isnumeric(request("groupid")) then
response.write "错误的参数！"
exit sub
end if
if request("groupaction")="yes" then
	dim GroupSetting
	GroupSetting=Request.Form("canview") & "," & Request.Form("canviewuserinfo") & "," & Request.Form("canviewpost") & "," & Request.Form("cannewpost") & "," & Request.Form("canreplymytopic") & "," & Request.Form("canreplytopic") & "," & Request.Form("canpostagree") & "," & Request.Form("canupload") & "," & Request.Form("canpostvote") & "," & Request.Form("canvote") & "," & Request.Form("caneditmytopic") & "," & Request.Form("candelmytopic") & "," & Request.Form("canmovemytopic") & "," & Request.Form("canclosemytopic") & "," & Request.Form("cansearch") & "," & Request.Form("canmailtopic") & "," & Request.Form("canmodify") & "," & Request.Form("cansmallpaper") & "," & Request.Form("candeltopic") & "," & Request.Form("canmovetopic") & "," & Request.Form("canclosetopic") & "," & Request.Form("cantoptopic") & "," & Request.Form("canawardtopic") & "," & Request.Form("canmodifytopic") & "," & Request.Form("canbesttopic") & "," & Request.Form("canAnnounce") & "," & Request.Form("canAdminAnnounce") & "," & Request.Form("canAdminPaper") & "," & Request.Form("canAdminUser") & "," & Request.Form("canDelUserTopic") & "," & Request.Form("canviewip") & "," & Request.Form("canadminip") & "," & Request.Form("cansendsms") & "," & Request.Form("Maxsendsms") & "," & Request.Form("Maxsmsbody") & "," & Request.Form("Maxsmsbox") & "," & Request.Form("canusetitle") & "," & Request.Form("canuseface") & "," & Request.Form("canusesign") & "," & Request.Form("canvieweven") & "," & Request.Form("canuploadnum") & "," & Request.Form("canviewbest") & "," & Request.Form("adminpermission") & "," & request.form("canaward") & "," & request.form("MaxUploadSize") & "," & request.form("canbatchtopic") & "," & request.form("smallpapermoney") & "," & request.form("postagreemoney") & "," & request.form("canadminfile") & "," & request.form("ba1") & "," & Request.Form("ba2") & "," & request.form("ba3") & "," & request.form("ba4") & "," & request.form("ba5") & "," & request.form("ba6") & "," & request.form("ba7")
	Set rs= Server.CreateObject("ADODB.Recordset")
	if request("isdefault")=1 then
	conn.execute("delete from BoardPermission where BoardID="&request("reBoardID")&" and GroupID="&request("GroupID"))
	else
	if request("pid")<>"" then
	sql="update BoardPermission set PSetting='"&GroupSetting&"' where pid="&request("pid")
	else
	sql="insert into BoardPermission (BoardID,GroupID,PSetting) values ("&request("reBoardID")&","&request("GroupID")&",'"&GroupSetting&"')"
	end if
	conn.execute(sql)
	end if
	set rs=nothing
	response.write "修改成功！返回<a href=?action=permission>论坛权限管理</a>"
else
Dim reGroupSetting,reBoardID,groupid
Dim Groupname,Boardname,founduserper
founduserper=false
if request("GroupID")<>"" then
set rs=conn.execute("select * from BoardPermission where boardid="&request("reBoardID")&" and GroupID="&request("GroupID"))
if rs.eof and rs.bof then
	founduserper=false
else
groupid=rs("groupid")
reGroupSetting=split(rs("PSetting"),",")
reBoardID=rs("boardid")
set rs=conn.execute("select title from UserGroups where usergroupid="&groupid)
groupname=rs("title")
founduserper=true
end if
if not founduserper then
set rs=conn.execute("select * from usergroups where usergroupid="&request("groupid"))
if rs.eof and rs.bof then
response.write "未找到该用户组！"
exit sub
end if
groupid=request("groupid")
reGroupSetting=split(rs("GroupSetting"),",")
reBoardID=request("reBoardID")
Groupname=rs("title")
end if
end if
set rs=conn.execute("select boardtype from board where boardid="&reBoardID)
Boardname=rs("boardtype")
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<FORM METHOD=POST ACTION="?action=editpermission">
<input type=hidden name="groupid" value="<%=groupid%>">
<input type=hidden name="reBoardID" value="<%=reBoardID%>">
<input type=hidden name="pID" value="<%
	dim ars
	set ars=conn.execute("select pid from BoardPermission where BoardID="&reBoardID&" and GroupID="&groupid)
	if not (ars.eof or ars.bof) then
		response.write ars("pid")
	end if
%>">
<tr> 
<th height="23" colspan="2" >编辑论坛用户组权限&nbsp;>> <%=boardname%>&nbsp;>> <%=groupname%></th>
</tr>
<tr> 
<td height="23" colspan="2" class=forumrow><input type=radio name="isdefault" value="1" <%if not founduserper then%>checked<%end if%>><B>使用用户组默认值</B> (注意: 这将删除任何之前所做的自定义设置)</td>
</tr>
<tr> 
<td height="23" colspan="2"  class=forumrow><input type=radio name="isdefault" value="0" <%if founduserper then%>checked<%end if%>><B>使用自定义设置</B>&nbsp;(选择自定义才能使以下设置生效) </td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝查看权限</th>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以浏览论坛</td>
<td height="23" width="40%" class=forumrow>是<input name="canview" type=radio value="1" <%if reGroupSetting(0)=1 then%>checked<%end if%>>&nbsp;否<input name="canview" type=radio value="0" <%if reGroupSetting(0)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以查看会员信息(包括其他会员的资料和会员列表)
</td>
<td height="23" width="40%" class=forumrow>是<input name="canviewuserinfo" type=radio value="1" <%if reGroupSetting(1)=1 then%>checked<%end if%>>&nbsp;否<input name="canviewuserinfo" type=radio value="0" <%if reGroupSetting(1)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以查看其他人发布的主题
</td>
<td height="23" width="40%" class=forumrow>是<input name="canviewpost" type=radio value="1" <%if reGroupSetting(2)=1 then%>checked<%end if%>>&nbsp;否<input name="canviewpost" type=radio value="0" <%if reGroupSetting(2)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以浏览精华帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canviewbest" type=radio value="1" <%if reGroupSetting(41)=1 then%>checked<%end if%>>&nbsp;否<input name="canviewbest" type=radio value="0" <%if reGroupSetting(41)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝<b>发帖权限</b></th>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以发布新主题</td>
<td height="23" width="40%" class=forumrow>是<input name="cannewpost" type=radio value="1" <%if reGroupSetting(3)=1 then%>checked<%end if%>>&nbsp;否<input name="cannewpost" type=radio value="0" <%if reGroupSetting(3)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以回复自己的主题
</td>
<td height="23" width="40%" class=forumrow>是<input name="canreplymytopic" type=radio value="1" <%if reGroupSetting(4)=1 then%>checked<%end if%>>&nbsp;否<input name="canreplymytopic" type=radio value="0" <%if reGroupSetting(4)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以回复其他人的主题
</td>
<td height="23" width="40%" class=forumrow>是<input name="canreplytopic" type=radio value="1" <%if reGroupSetting(5)=1 then%>checked<%end if%>>&nbsp;否<input name="canreplytopic" type=radio value="0" <%if reGroupSetting(5)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以在论坛允许评分的时候参与评分(鲜花和鸡蛋)?
</td>
<td height="23" width="40%" class=forumrow>是<input name="canpostagree" type=radio value="1" <%if reGroupSetting(6)=1 then%>checked<%end if%>>&nbsp;否<input name="canpostagree" type=radio value="0" <%if reGroupSetting(6)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>参与评分所需金钱
</td>
<td height="23" width="40%" class=Forumrow><input name="postagreemoney" type=text size=4 value="<%=reGroupSetting(47)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以上传附件
</td>
<td height="23" width="40%" class=forumrow>是<input name="canupload" type=radio value="1" <%if reGroupSetting(7)=1 then%>checked<%end if%>>&nbsp;否<input name="canupload" type=radio value="0" <%if reGroupSetting(7)=0 then%>checked<%end if%>>
&nbsp;发帖可以上传<input name="canupload" type=radio value="2" <%if reGroupSetting(7)=2 then%>checked<%end if%>>&nbsp;回复可以上传<input name="canupload" type=radio value="3" <%if reGroupSetting(7)=3 then%>checked<%end if%>>
</td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>一次最多上传文件个数
</td>
<td height="23" width="40%" class=Forumrow><input name="canuploadnum" type=text size=4 value="<%=reGroupSetting(40)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>一天最多上传文件个数
</td>
<td height="23" width="40%" class=Forumrow><input name="ba2" type=text size=4 value="<%=reGroupSetting(50)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>上传文件大小限制
</td>
<td height="23" width="40%" class=Forumrow><input name="MaxUploadSize" type=text size=4 value="<%=reGroupSetting(44)%>"> KB</td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以发布新投票</td>
<td height="23" width="40%" class=forumrow>是<input name="canpostvote" type=radio value="1" <%if reGroupSetting(8)=1 then%>checked<%end if%>>&nbsp;否<input name="canpostvote" type=radio value="0" <%if reGroupSetting(8)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以参与投票</td>
<td height="23" width="40%" class=forumrow>是<input name="canvote" type=radio value="1" <%if reGroupSetting(9)=1 then%>checked<%end if%>>&nbsp;否<input name="canvote" type=radio value="0" <%if reGroupSetting(9)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以发布小字报</td>
<td height="23" width="40%" class=Forumrow>是<input name="cansmallpaper" type=radio value="1"  <%if reGroupSetting(17)=1 then%>checked<%end if%>>&nbsp;否<input name="cansmallpaper" type=radio value="0" <%if reGroupSetting(17)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>发布小字报所需金钱</td>
<td height="23" width="40%" class=Forumrow><input name="smallpapermoney" type=text value="<%=reGroupSetting(46)%>" size=4></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝<b>帖子/主题编辑权限</b></th>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以编辑自己的帖子
</td>
<td height="23" width="40%" class=forumrow>是<input name="caneditmytopic" type=radio value="1" <%if reGroupSetting(10)=1 then%>checked<%end if%>>&nbsp;否<input name="caneditmytopic" type=radio value="0" <%if reGroupSetting(10)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以删除自己的帖子
</td>
<td height="23" width="40%" class=forumrow>是<input name="candelmytopic" type=radio value="1" <%if reGroupSetting(11)=1 then%>checked<%end if%>>&nbsp;否<input name="candelmytopic" type=radio value="0" <%if reGroupSetting(11)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以移动自己的帖子到其他论坛
</td>
<td height="23" width="40%" class=forumrow>是<input name="canmovemytopic" type=radio value="1" <%if reGroupSetting(12)=1 then%>checked<%end if%>>&nbsp;否<input name="canmovemytopic" type=radio value="0" <%if reGroupSetting(12)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以打开/关闭自己发布的主题
</td>
<td height="23" width="40%" class=forumrow>是<input name="canclosemytopic" type=radio value="1" <%if reGroupSetting(13)=1 then%>checked<%end if%>>&nbsp;否<input name="canclosemytopic" type=radio value="0" <%if reGroupSetting(13)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝<b>其他权限</b></th>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以搜索论坛
</td>
<td height="23" width="40%" class=forumrow>是<input name="cansearch" type=radio value="1" <%if reGroupSetting(14)=1 then%>checked<%end if%>>&nbsp;否<input name="cansearch" type=radio value="0" <%if reGroupSetting(14)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以使用'发送本页给好友'功能
</td>
<td height="23" width="40%" class=forumrow>是<input name="canmailtopic" type=radio value="1" <%if reGroupSetting(15)=1 then%>checked<%end if%>>&nbsp;否<input name="canmailtopic" type=radio value="0" <%if reGroupSetting(15)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以修改个人资料
</td>
<td height="23" width="40%" class=forumrow>是<input name="canmodify" type=radio value="1" <%if reGroupSetting(16)=1 then%>checked<%end if%>>&nbsp;否<input name="canmodify" type=radio value="0" <%if reGroupSetting(16)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以浏览论坛事件
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canvieweven" type=radio value="1"  <%if reGroupSetting(39)=1 then%>checked<%end if%>>&nbsp;否<input name="canvieweven" type=radio value="0" <%if reGroupSetting(39)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可浏览论坛展区的权限
</td>
<td height="23" width="40%" class=Forumrow>是<input name="ba1" type=radio value="1"  <%if reGroupSetting(49)=1 then%>checked<%end if%>>&nbsp;否<input name="ba1" type=radio value="0" <%if reGroupSetting(49)=0 then%>checked<%end if%>></td>
</tr>
<input type=hidden value=0 name="ba4">
<input type=hidden value=0 name="ba5">
<input type=hidden value=0 name="ba6">
<input type=hidden value=0 name="ba7">
<tr> 
<th height="23" colspan="2" align=left>＝＝管理权限</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以删除其它人帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="candeltopic" type=radio value="1" <%if reGroupSetting(18)=1 then%>checked<%end if%>>&nbsp;否<input name="candeltopic" type=radio value="0"  <%if reGroupSetting(18)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以移动其它人帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canmovetopic" type=radio value="1" <%if reGroupSetting(19)=1 then%>checked<%end if%>>&nbsp;否<input name="canmovetopic" type=radio value="0"  <%if reGroupSetting(19)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以打开/关闭其它人帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canclosetopic" type=radio value="1" <%if reGroupSetting(20)=1 then%>checked<%end if%>>&nbsp;否<input name="canclosetopic" type=radio value="0"  <%if reGroupSetting(20)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以固顶/解除固顶帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="cantoptopic" type=radio value="1" <%if reGroupSetting(21)=1 then%>checked<%end if%>>&nbsp;否<input name="cantoptopic" type=radio value="0"  <%if reGroupSetting(21)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以进行帖子总固顶操作
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canusesign" type=radio value="1"  <%if reGroupSetting(38)=1 then%>checked<%end if%>>&nbsp;否<input name="canusesign" type=radio value="0" <%if reGroupSetting(38)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以奖励/惩罚发帖用户
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canawardtopic" type=radio value="1" <%if reGroupSetting(22)=1 then%>checked<%end if%>>&nbsp;否<input name="canawardtopic" type=radio value="0"  <%if reGroupSetting(22)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以奖励/惩罚用户
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canaward" type=radio value="1" <%if reGroupSetting(43)=1 then%>checked<%end if%>>&nbsp;否<input name="canaward" type=radio value="0"  <%if reGroupSetting(43)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以编辑其它人帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canmodifytopic" type=radio value="1" <%if reGroupSetting(23)=1 then%>checked<%end if%>>&nbsp;否<input name="canmodifytopic" type=radio value="0" <%if reGroupSetting(23)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以加入/解除精华帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canbesttopic" type=radio value="1" <%if reGroupSetting(24)=1 then%>checked<%end if%>>&nbsp;否<input name="canbesttopic" type=radio value="0"  <%if reGroupSetting(24)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以发布公告
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canAnnounce" type=radio value="1" <%if reGroupSetting(25)=1 then%>checked<%end if%>>&nbsp;否<input name="canAnnounce" type=radio value="0"  <%if reGroupSetting(25)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以管理公告
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canAdminAnnounce" type=radio value="1" <%if reGroupSetting(26)=1 then%>checked<%end if%>>&nbsp;否<input name="canAdminAnnounce" type=radio value="0"  <%if reGroupSetting(26)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以管理小字报
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canAdminPaper" type=radio value="1" <%if reGroupSetting(27)=1 then%>checked<%end if%>>&nbsp;否<input name="canAdminPaper" type=radio value="0"  <%if reGroupSetting(27)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以锁定/屏蔽/解除锁定用户
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canAdminUser" type=radio value="1" <%if reGroupSetting(28)=1 then%>checked<%end if%>>&nbsp;否<input name="canAdminUser" type=radio value="0"  <%if reGroupSetting(28)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以删除用户1－10天内所发帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canDelUserTopic" type=radio value="1" <%if reGroupSetting(29)=1 then%>checked<%end if%>>&nbsp;否<input name="canDelUserTopic" type=radio value="0"  <%if reGroupSetting(29)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以查看来访IP及来源
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canviewip" type=radio value="1" <%if reGroupSetting(30)=1 then%>checked<%end if%>>&nbsp;否<input name="canviewip" type=radio value="0"  <%if reGroupSetting(30)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以限定IP来访
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canadminip" type=radio value="1" <%if reGroupSetting(31)=1 then%>checked<%end if%>>&nbsp;否<input name="canadminip" type=radio value="0"  <%if reGroupSetting(31)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以管理用户权限
</td>
<td height="23" width="40%" class=Forumrow>是<input name="adminpermission" type=radio value="1" <%if reGroupSetting(42)=1 then%>checked<%end if%>>&nbsp;否<input name="adminpermission" type=radio value="0"  <%if reGroupSetting(42)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以批量删除帖子（前台）
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canbatchtopic" type=radio value="1" <%if reGroupSetting(45)=1 then%>checked<%end if%>>&nbsp;否<input name="canbatchtopic" type=radio value="0"  <%if reGroupSetting(45)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>是否有审核帖子的权限
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canusetitle" type=radio value="1" <%if reGroupSetting(36)=1 then%>checked<%end if%>>&nbsp;否<input name="canusetitle" type=radio value="0" <%if reGroupSetting(36)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>是否有进入隐含论坛的权限<BR>在用户组中开放则可进入所有隐含论坛<BR>在论坛权限管理中设置可进入设置的版面<BR>在用户权限管理中设置可进入设置的版面
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canuseface" type=radio value="1"  <%if reGroupSetting(37)=1 then%>checked<%end if%>>&nbsp;否<input name="canuseface" type=radio value="0" <%if reGroupSetting(37)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>有论坛文件管理权限
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canadminfile" type=radio value="1" <%if reGroupSetting(48)=1 then%>checked<%end if%>>&nbsp;否<input name="canadminfile" type=radio value="0" <%if reGroupSetting(48)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>有帖子主题颜色权限
</td>
<td height="23" width="40%" class=Forumrow>是<input name="ba3" type=radio value="1" <%if reGroupSetting(51)=1 then%>checked<%end if%>>&nbsp;否<input name="ba3" type=radio value="0" <%if reGroupSetting(51)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝短信权限</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以发送短信
</td>
<td height="23" width="40%" class=Forumrow>是<input name="cansendsms" type=radio value="1"  <%if reGroupSetting(32)=1 then%>checked<%end if%>>&nbsp;否<input name="cansendsms" type=radio value="0" <%if reGroupSetting(32)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>最多发送用户
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsendsms" size=5 type=text value="<%=reGroupSetting(33)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>短信内容大小限制
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsmsbody" size=5 type=text value="<%=reGroupSetting(34)%>"> byte</td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>信箱大小限制
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsmsbox" size=5 type=text value="<%=reGroupSetting(35)%>"> KB</td>
</tr>
<tr> 
<td height="23" width="60%" class=forumrow>
</td>
<td height="23" width="40%" class=forumrow><input type="submit" name="submit" value="提 交"></td>
</tr>
<input type=hidden value="yes" name="groupaction">
</FORM>
</table>
<%
end if
end sub

sub cache_board()
'cache版面数据
myCache.name="BoardJumpList"
Dim BoardJumpList
set rs=conn.execute("select boardid,boardtype,depth from board order by rootid,orders")
do while not rs.EOF
BoardJumpList = BoardJumpList & "<option value=""list.asp?boardid="&rs(0)&""" "
BoardJumpList = BoardJumpList & ">"
select case rs(2)
case 0
BoardJumpList = BoardJumpList & "╋"
case 1
BoardJumpList = BoardJumpList & "&nbsp;&nbsp;├"
end select
if rs(2)>1 then
for i=2 to rs(2)
	BoardJumpList = BoardJumpList & "&nbsp;&nbsp;│"
next
BoardJumpList = BoardJumpList & "&nbsp;&nbsp;├"
end if
BoardJumpList = BoardJumpList & rs(1)&"</option>"
rs.MoveNext
loop
myCache.add BoardJumpList,dateadd("n",9999,now)
set rs=nothing
'end cache
end sub

sub RestoreBoard()
'按照目前的排序循环i数值更新rootid
'还原所有版面的depth,orders,parentid,parentstr,child为0
i=0
set rs=conn.execute("select boardid from board order by rootid,orders")
do while not rs.eof
i=i+1
conn.execute("update board set rootid="&i&",depth=0,orders=0,ParentID=0,ParentStr='0',child=0 where boardid="&rs(0))
rs.movenext
loop
set rs=nothing
response.write "复位成功，请返回做论坛归属设置。"
call cache_board()
end sub
%>