<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_down_conn.asp"-->
<html>
<head>
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0">
<%dim str
if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	call dvbbs_error()
else
	if Request("action") = "unite" then
		call unite()
	elseif Request("action") = "npwd" then
		call nodownpwd()
	elseif Request("action") = "deltime" then
		call deltime()
	elseif Request("action") = "deluser" then
		call deluser()
	elseif Request("action") = "dpath" then
		call dpath()
	elseif Request("action") = "show" then
		call show()
	elseif Request("action") = "show2" then
		call show2()
	elseif Request("action") = "save" then
		call save()
	elseif Request("action") = "vipsave" then
		call vipsave()
	elseif Request("action") = "upset" then
		call upset()
	elseif Request("action") = "retime" then
		call retime()
	elseif Request("action") = "renfile" then
		call renfile()
	else
		call boardinfo()
	end if
end if
conndown.close
set conndown=nothing%>
</html>

<%sub boardinfo()%>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<form action=?action=unite method=post>
		<tr>
			<th>合并子分类数据</th>
		</tr>
		<tr>
			<td class=forumrow>注意事项：在下面，您将看到目前所有的论坛列表。您可以将一个子分类和另外一个子分类进行合并，合并后两个子分类所有软件转入合并后的子分类。</td>
		</tr>
		<tr>
			<td class=forumrow><%
				dim NClassString
				set rs = server.CreateObject ("Adodb.recordset")
				sql="select * from DNclass"
				rs.open sql,conndown,1,1
				if rs.eof and rs.bof then
					NClassString="没有子分类"
				else
					NClassString=""
					do while not rs.eof
						NClassString=NClassString&"<option value="&rs("nclassid")&">"&rs("nclass")&"</option>"
						rs.movenext
					loop
					response.write "将子分类&nbsp;"
					response.write "<select name=oldboard size=1>"
					response.write NClassString
					response.write "</select>&nbsp;&nbsp;"
					response.write "合并到&nbsp;"
					response.write "<select name=newboard size=1>"
					response.write NClassString
					response.write "</select>&nbsp;&nbsp;"
				end if
				rs.close
				set rs=nothing
				response.write "<input type=submit name=Submit value=执行>"
			%></td>
		</tr>
		</form>
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<tr>
			<th>防 盗 链</th>
		</tr>
		<tr>
			<td class=forumrow>防盗链密码可以有效的防止别人从外部盗链您的下载地址，请定期修改<br>最有效的防盗链手段还是定期更新下载地址</td>
		</tr>
		<%set rs = server.CreateObject ("Adodb.recordset")
		sql="select * from notdown"
		rs.open sql,conndown,1,1	%>
		<form action=?action=npwd method=post>
		<tr>
			<td class=forumrow><b>防盗链密码：</b><input maxLength=20 name=downpwd size=15 value="<%=rs("nodown")%>">&nbsp;&nbsp;<input type=submit name=submit value="修 改"></td>
		</tr>
		</form>
		<tr>
			<td class=forumrow>定期修改您的软件存放目录的名称，可以有效的防止别人盗链您的软件<br><font color="#FF0000">说明：如果您的空间支持FSO功能，那么您填写完文件夹名，点击修改后即可完成目录名的修改<br>如果您的空间不支持FSO，那么请您修改完后，自已手工修改原下载目录名，使其和您修改的目录名一致<br></font><font color="#FF0000">填写规则</font>：您所填写的文件夹必须和index.asp文件在同一层目录，并且不要填写子目录名。<br><font color="#FF0000">正确填写如</font>：uploadimages</td>
		</tr>
		<form action=?action=dpath method=post>
		<tr>
			<td class=forumrow><b>软件存放目录：</b><input maxLength=20 name=downpath size=15 value="<%=rs("downpath")%>">&nbsp;&nbsp;<input type=submit name=submit value="修 改"></td>
		</tr>
		</form>
		<%rs.close
		set rs=nothing%>
		<form action=?action=renfile method=post>
		<tr>
			<td class=forumrow><b>文件随机改名：</b>&nbsp;<input type=submit name=submit value="改 名"></td>
		</tr>
		</form>
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<form action=?action=deltime method=post>
		<tr>
			<th valign=middle colspan=2 height=23 align=left>删除指定日期内事件</th>
		</tr>
		<tr>
			<td valign=middle width=40% class=forumrow>删除多少天前的事件(填写数字)</td><td class=forumrow><input name="TimeLimited" value=0 size=30>&nbsp;<input type=submit name="submit" value="提 交"></td>
		</tr>
		<tr>
			<td valign=middle width=40%  class=forumrow>事件类型</td>
			<td class=forumrow><select name="deltype" size=1><option value="1">管理员管理日志</option><option value="2">付费用户下载日志</option><option value="3">付费用户管理日志</option></select></td>
		</tr>
		</form>
		<form action=?action=deluser method=post>
		<tr>
			<th valign=middle colspan=2 height=23 align=left>删除某用户的所有事件</th>
		</tr>
		<tr>
			<td valign=middle width=40%  class=forumrow>请输入用户名</td>
			<td class=forumrow><input type=text name="username" size=30>&nbsp;<input type=submit name="submit" value="提 交"></td>
		</tr>
		<tr>
			<td valign=middle width=40%  class=forumrow>事件类型</td>
			<td class=forumrow><select name="deltype" size=1><option value="1">付费用户下载日志</option><option value="2">付费用户管理日志</option></select></td>
		</tr>
		</form>
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<%dim ShowSet,upsizemax,uppicsizemax
		set rs = server.CreateObject ("Adodb.recordset")
		sql="select * from showpage"
		rs.open sql,conndown,1,1
		showset=split(rs("showset"),",")
		upsizemax=int(int(rs("upsizemax"))/1024)
		uppicsizemax=int(int(rs("uppicsizemax"))/1024)%>
		<form action=?action=upset method=post>
		<tr>
			<th valign=middle height=23 align=left width="100%" colspan="2">上传属性设定</th>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="18%">上传<b>文件</b>类型(请用|分隔)：</td> 
	  	<td valign=middle class=forumrow align=left width="57%"><input type=text name="uptype" size=74 value="<%=rs("uptype")%>"></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="18%">上传<b>文件大小</b>设定：</td>
			<td valign=middle class=forumrow align=left width="57%"><input type=text name="upsizemin" value="<%=rs("upsizemin")%>" size=13><b>K</b>&nbsp; &lt; 上传文件大小 &lt; <input type=text name="upsizemax" value="<%=upsizemax%>" size=13><b>K</b></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="18%">上传<b>软件图片</b>类型(请用|分隔)：</td>
			<td valign=middle class=forumrow align=left width="57%"><input type=text name="uppictype" size=74 value="<%=rs("uppictype")%>"></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="18%">上传<b>软件图片大小</b>设定：</td>
			<td valign=middle class=forumrow align=left width="57%"><input type=text name="uppicsizemin" value="<%=rs("uppicsizemin")%>" size=13><b>K</b>&nbsp;&lt; 上传软件图片大小 &lt; <input type=text name="uppicsizemax" value="<%=uppicsizemax%>" size=13><b>K</b></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=center colspan=2><input type=submit name="submit" value="提 交"></td>
		</tr>
		</form>
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<form action=?action=vipsave method=post>
		<tr>
			<th valign=middle height=23 align=left width="100%">VIP/普通版切换</th>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="100%"><font color="#FF0000">重要说明：</font>如果您的认坛没有安装VIP插件，那么您只能使用普通版，如果选择VIP版，将会出现错误。<br>如果您的论坛安装了VIP插件，建议您使用VIP版本<br>区别：VIP版本将会比普通版多一个VIP会员下载，这个功能可以使下载仅限于论坛VIP会员<br>注意：<font color="#FF0000">如果您在使用了VIP版一段时间后，切换到普通版，您论坛中所有的仅供VIP下载的软件将变成仅限会员下载</font>，请仔细考虑后操作</td>
		</tr>
		<tr>
			<td valign=middle class=forumrow width="100%" align="center">VIP会员版&nbsp;<input type=radio name="vipshow" value="2" <%if rs("vipshow")=2 then%>checked<%end if%>>&nbsp;&nbsp;&nbsp;&nbsp;普通版本&nbsp;<input type=radio name="vipshow" value="1" <%if rs("vipshow")=1 then%>checked<%end if%>></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow width="100%" align="center"><input type=submit name="submit" value="提 交"></td>
		</tr>
		</form>
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<form action=?action=retime method=post>
		<tr>
			<th valign=middle height=23 align=left width="100%">页面刷新时间限定</th>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="100%">说明：用户在<b>软件详细资料页面</b>刷新时间不得超过此设定时间，设定时间过长会影响用户浏览，时间太短可能会遇到恶意刷新。请仔细设定</td> 
		</tr>
		<tr>
			<td valign=middle class=forumrow align=center width="100%"><input type=text name="retime" value="<%=rs("retime")%>" size=13><b>秒</b>&nbsp;&nbsp;&nbsp;&nbsp;<input type=submit name="submit" value="提 交"></td>
		</tr>
		</form>
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<form action=?action=show method=post>
		<tr>
			<th valign=middle colspan=4 height=23 align=left width="100%">页面显示参数设定</th>
		</tr>
		<tr>
			<td valign=middle class=forumrow width="25%">当日下载TOP数：</td>
			<td valign=middle class=forumrow width="175"><input type=text name="daytop" size=15 value="<%=rs("daytop")%>"></td>
			<td valign=middle class=forumrow width="163">本周下载TOP数：</td>
			<td valign=middle class=forumrow width="189"><input type=text name="weektop" size=15 value="<%=rs("weektop")%>"></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow width="25%">累计下载TOP数：</td>
			<td valign=middle class=forumrow width="175"><input type=text name="alltop" size=15 value="<%=rs("alltop")%>"></td>
			<td valign=middle class=forumrow width="163">热门推荐TOP数：</td>
			<td valign=middle class=forumrow width="189"><input type=text name="hotnum" size=15 value="<%=rs("hotnum")%>"></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow width="25%">固顶推荐TOP数：</td>
			<td valign=middle class=forumrow width="175"><input type=text name="gudinnum" size=15 value="<%=rs("gudinnum")%>"></td>
			<td valign=middle class=forumrow width="163">软件查找结果每页显示：</td>
			<td valign=middle class=forumrow width="189"><input type=text name="sonum" size=15 value="<%=rs("sonum")%>"></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow width="25%">软件列表中每页数目：</td>
			<td valign=middle class=forumrow width="175"><input type=text name="listnum" size=15 value="<%=rs("listnum")%>"></td>
			<td valign=middle class=forumrow width="163">新进软件列表数：</td>
			<td valign=middle class=forumrow width="189"><input type=text name="newnum" size=15 value="<%=rs("newnum")%>"></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow width="25%">修改删除页面每页软件数目：</td>
			<td valign=middle class=forumrow width="175"><input type=text name="editnum" size=15 value="<%=rs("editnum")%>"></td>
			<td valign=middle class=forumrow width="163">各类事件页面每页事件数目：</td>
			<td valign=middle class=forumrow width="189"><input type=text name="eventnum" size=15 value="<%=rs("eventnum")%>"></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow colspan=4 align=center><input type=submit name="submit" value="提 交"></td>
		</tr>
		</form>
		<%rs.close
		set rs=nothing%>
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<form action=?action=show2 method=post>
		<tr>
			<th valign=middle colspan=6 height=23 align=left width="100%">索引显示设定</th>
		</tr>
		<tr>
			<td valign=middle class=forumrow width="20%" align="center"></td>
			<td valign=middle class=forumrow width="16%" align="center">固顶推荐</td>
			<td valign=middle class=forumrow width="16%" align="center">热门推荐</td>
			<td valign=middle class=forumrow width="16%" align="center">今日下载</td>
			<td valign=middle class=forumrow width="16%" align="center">本周下载</td>
			<td valign=middle class=forumrow width="16%" align="center">累计下载</td>
		</tr>
		<tr>
			<td valign=middle class=forumrow width="20%" nowrap>下载中心首页(z_down_default.asp)</td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="allhot1"   value="1" <%if showset(0)=1 then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="hot1" value="1" <%if showset(1)=1   then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="daytop1" value="1" <%if showset(2)=1 then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="weektop1"   value="1" <%if showset(3)=1 then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="alltop1" value="1" <%if showset(4)=1   then%>checked<%end if%>></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow width="20%" nowrap>软件分类页面(z_down_index.asp)</td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="allhot2"   value="1" <%if showset(5)=1 then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="hot2" value="1" <%if showset(6)=1 then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="daytop2" value="1" <%if showset(7)=1 then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="weektop2"   value="1" <%if showset(8)=1 then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="alltop2" value="1" <%if showset(9)=1   then%>checked<%end if%>></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow width="20%" nowrap>软件详细资料页面(z_down_list.asp)</td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="allhot3" value="1" <%if showset(10)=1 then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="hot3" value="1" <%if showset(11)=1 then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="daytop3" value="1" <%if showset(12)=1 then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="weektop3" value="1" <%if showset(13)=1 then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="alltop3" value="1" <%if showset(14)=1 then%>checked<%end if%>></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow colspan=6 align=center><input type=submit name="submit" value="修 改"></td>
		</tr>
		</form>
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<tr>
			<th valign=middle colspan=2 height=23 align=left width="100%">上传文件管理（需要FSO支持）</th>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="25%"> 说明： </td>
			<td valign=middle class=forumrow align=left width="75%">你的空间必须要支持FSO才能使用本功能</td>
		</tr> 
		<tr>
			<form name="form1" method="get" action="z_admin_down_uploadlist.asp">
			<td valign=middle class=forumrow colspan=2 align=center><INPUT type=submit value='进 行 管 理' name=submit></td>
			</form>
		</tr> 
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<form action=?action=save method=post>
		<tr>
			<th valign=middle height=23 align=left width="100%">在线数据库操作</th>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="100%">指令：<Input type="text" name="SQL_Statement" Size=60> <Input type="Submit" Value="送出"></td> 
		</tr>
		</form>
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<form action=z_admin_down_rb.asp?action=backupdata method=post>
		<tr>
			<th valign=middle colspan=2 height=23 align=left width="100%">下载数据库备份（需要FSO支持）</th>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="25%">当前数据库路径：&nbsp;</td>
			<td valign=middle class=forumrow align=left width="75%"><input type=text size=24 name=DBpath value="<%=dbdown%>"> 请正确添写您当前使用的数据库路径！</td>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="25%">备份数据库目录：</td>
			<td valign=middle class=forumrow align=left width="75%"><input type=text size=24 name=bkfolder value=Databackup> 如果目录不存在，程序将自动创建！</td>
		</tr> 
		<tr>
			<td valign=middle class=forumrow align=left width="25%">备份数据库名称：</td>
			<td valign=middle class=forumrow align=left width="75%"><input type=text size=24 name=bkDBname value=downbackup.mdb> 如果备份目录有该文件，将覆盖，如果没有，程序将自动创建！</td>
		</tr> 
	  <tr>
	  	<td valign=middle class=forumrow colspan=2 align=center><input type=submit value="开始备份"></td>
	  </tr>
		</form>
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<form action=z_admin_down_rb.asp?action=restoredata method=post>
		<tr>
			<th valign=middle colspan=2 height=23 align=left width="100%">下载数据库恢复（需要FSO支持）</th>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="25%"> 备份数据库路径：</td>
			<td valign=middle class=forumrow align=left width="75%"><input type=text size=37 name=dbpath value="Databackup/downbackup.mdb">&nbsp;如果文件不存在，将不能恢复</td>
		</tr> 
		<tr>
			<td valign=middle class=forumrow align=left width="25%"> 目标数据库路径：</td>
			<td valign=middle class=forumrow align=left width="75%"><input type=text size=37 name=backpath value="<%=dbdown%>">&nbsp;填写您当前使用的数据库路径！</td>
		</tr> 
	  <tr><td valign=middle class=forumrow colspan=2 align=center><input type=submit value="开始恢复"></td></tr>
	</form>
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<form action=z_admin_down_rb.asp?action=comp method=post>
		<tr>
			<th valign=middle colspan=2 height=23 align=left width="100%">压缩下载数据库（需要FSO支持）</th>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="100%" colspan="2"> <b>注意：</b><br>输入数据库所在相对路径,并且输入数据库名称（正在使用中数据库不能压缩，请选择备份数据库进行压缩操作）</td>
		</tr> 
		<tr>
			<td valign=middle class=forumrow align=left width="25%">压缩数据库：</td>
			<td valign=middle class=forumrow align=left width="75%"><input type="text" name="dbpath" value=Databackup/downbackup.mdb size="39">&nbsp;<input type="submit" value="开始压缩"></td>
		</tr> 
	  <tr>
	  	<td valign=middle class=forumrow align=left width="100%" colspan="2"><input type="checkbox" name="boolIs97" value="True"> 如果使用 Access 97 数据库请选择(默认为 Access 2000 数据库)</td>
		</tr>
		</form>
	</table>
	<br>
<%end sub

sub unite()
	dim newboard,newclassid,oldboard
	if cint(request("newboard"))=cint(request("oldboard")) then
		response.write "请不要在相同子分类内进行操作。"
	else
		newboard=cint(request("newboard"))
		oldboard=cint(request("oldboard"))
		set rs = server.CreateObject ("Adodb.recordset")
		sql="select * from DNclass where nclassid="&newboard&""
		rs.open sql,conndown,1,1
		newclassid=rs("classid")
		rs.close
		set rs=nothing
		conndown.execute("update download set nclassid="&newboard&",classid="&newclassid&" where nclassid="&oldboard)
		response.write "合并成功，已经将被合并子分类所有数据转入您所合并子分类。"
	end if
end sub

sub nodownpwd()
	dim nopwd
	if request("downpwd")="" then
		response.write "防盗链密码不能为空。"
	else
		nopwd=request("downpwd")
		set rs = server.CreateObject ("Adodb.recordset")
		sql="select * from notdown"
		rs.open sql,conndown,1,3
		rs("nodown")=nopwd
		rs.update
		rs.close
		sql="select * from events"
		rs.open sql,conndown,1,3
		rs.addnew
		rs("event")=membername&"修改了防盗链密码！"
		rs("addtime")=date()
		rs.update
		rs.close
		set rs=nothing
		response.write "修改成功。"
	end if
end sub

sub dpath()
	dim downpath,downpathold,fsdir,fsdir1,fs,folder1,folder2
	downpath=request("downpath")
	if InStr(downpath,"/")<>0 or downpath="" then
		response.write("您输入的文件目录名不正确！请仔细阅读说明")
	else
		set rs = server.CreateObject ("Adodb.recordset")
		sql="select * from notdown"
		rs.open sql,conndown,1,3
		downpathold=rs("downpath")
		if downpath=downpathold then
			response.write("您输入的文件目录名与原文件夹名相同！无法正常操作！请仔细阅读说明")
		else
			rs("downpath")=downpath
			rs.update
			response.write("数据库中下载目录名修改成功！<br>")
			fsdir=Server.MapPath("index.asp")	
			fsdir1=replace(fsdir,"\index.asp","\")
			set fs=createobject("scripting.filesystemobject") 
			folder1=fsdir1&"\"&downpathold
			folder2=fsdir1&"\"&downpath
			If not fs.folderexists(folder1) then
				response.write fsDir&"原文件夹不存在,请重新输入文件夹名！<br>"
			elseIf fs.folderexists(folder2) then
				response.write "目标文件夹<font color=red>已经存在</font>,请选择其他文件夹<br>"
			else 
				fs.MoveFolder folder1,folder2
				If not fs.folderexists(folder1) and fs.folderexists(folder2) then
					response.write "文件目录名修改<font color=red>成功！</font><br>"
				Else
					response.write "文件目录名修改<font color=red>失败！</font><br>"
				End If
			end if
		end if
		rs.close
		set rs=nothing
	end if
end sub

sub deltime()
	dim dtime,deltype
	dtime=request("TimeLimited")
	if request("deltype")=1 then
		deltype="events"
	elseif request("deltype")=2 then
		deltype="downevent"
	elseif request("deltype")=3 then
		deltype="userevent"
	end if 
	if not isInteger(dtime) then
		response.write "输入的日期需要为整数。"
	else
		conndown.execute("delete from "&deltype&" where addtime>now()-"&dtime)
		response.write "删除成功。"
	end if
end sub 

sub deluser()
	dim username,deltype
	username=request("username")
	if request("deltype")=1 then
		deltype="downevent"
	elseif request("deltype")=2 then
		deltype="userevent"
	end if 
	set rs=conndown.execute("select id from [user] where username='"&username&"'")
	if rs.eof and rs.bof then
		response.write "此用户不是付费用户。"
	else
		conndown.execute("delete from "&deltype&" where username='"&username&"'")
		response.write "删除成功。"
	end if
	rs.close
end sub 

sub upset()
	dim upsizemax,uppicsizemax
	if request("uptype")="" or request("upsizemin")="" or request("upsizemax")="" or request("uppictype")="" or request("uppicsizemin")="" or request("uppicsizemax")="" then
		response.write "输入的内容不能为空！"
	elseif not isInteger(request("upsizemin")) or  not isInteger(request("upsizemax")) or not isInteger(request("uppicsizemin")) or  not isInteger(request("uppicsizemax")) then
		response.write "输入的上传文件大小必须为整数！"
	else
		upsizemax= int(request("upsizemax")*1024)
		uppicsizemax=int(request("uppicsizemax")*1024)
		set rs = server.CreateObject ("Adodb.recordset")
		sql="select * from showpage"
		rs.open sql,conndown,1,3
		rs("uptype")= request("uptype")
		rs("upsizemin")=int(request("upsizemin"))
		rs("upsizemax")=upsizemax
		rs("uppictype")= request("uppictype")
		rs("uppicsizemin")=int(request("uppicsizemin"))
		rs("uppicsizemax")=uppicsizemax
		rs.update
		rs.close
		set rs=nothing
		response.write "修改成功！"
	end if
end sub

sub vipsave()
	if request("vipshow")=""  then
		response.write "输入的数据不能为空！"
	else
		set rs = server.CreateObject ("Adodb.recordset")
		sql="select * from showpage"
		rs.open sql,conndown,1,3
		if request("vipshow")=2 or (request("vipshow")=1 and rs("vipshow")=1) then
			rs("vipshow")=request("vipshow")
			rs.update
			response.write("修改成功！")
		elseif request("vipshow")=1 and rs("vipshow")=2 then
			conndown.execute"update [download] set downshow=2 where downshow=3" 
			rs("vipshow")=request("vipshow")
			rs.update
			response.write("修改成功！")
		end if
		rs.close
		set rs=nothing
	end if
end sub

sub retime()
	if request("retime")="" then
		response.write "输入的内容不能为空！"
	elseif not isInteger(request("retime")) then
		response.write "输入的时间必须为整数！"
	else
		set rs = server.CreateObject ("Adodb.recordset")
		sql="select * from showpage"
		rs.open sql,conndown,1,3
		rs("retime")=request("retime")
		rs.update
		rs.close
		set rs=nothing
		response.write "修改成功！"
	end if
end sub

sub show()
	if request("daytop")="" or request("weektop")="" or request("alltop")="" or request("listnum")="" or request("editnum")="" or request("eventnum")="" or request("hotnum")="" or request("newnum")="" or request("gudinnum")="" then
		response.write "输入的数据不能为空！"
	elseif not isInteger(request("daytop")) or not isInteger(request("weektop")) or not isInteger(request("alltop")) or not isInteger(request("listnum")) or not isInteger(request("editnum")) or not isInteger(request("eventnum")) or not isInteger(request("newnum")) or not isInteger(request("hotnum")) or not isInteger(request("gudinnum")) then
		response.write "输入的数据必须为整数！"
	else
		set rs = server.CreateObject ("Adodb.recordset")
		sql="select * from showpage"
		rs.open sql,conndown,1,3
		rs("daytop")=request("daytop")
		rs("weektop")=request("weektop")
		rs("alltop")=request("alltop")
		rs("listnum")=request("listnum")
		rs("editnum")=request("editnum")
		rs("eventnum")=request("eventnum")
		rs("newnum")=request("newnum")
		rs("hotnum")=request("hotnum")
		rs("gudinnum")=request("gudinnum")
		rs("sonum")=request("sonum")
		rs.update
		rs.close
		response.write "修改成功！"
	end if
end sub

sub show2()
	dim allhot1,allhot2,allhot3,hot1,hot2,hot3,daytop1,daytop2,daytop3,weektop1,weektop2,weektop3,alltop1,alltop2,alltop3
	if request("allhot1")=1 then
		allhot1=1
	else
		allhot1=0
	end if
	if request("hot1")=1 then
		hot1=1
	else
		hot1=0
	end if
	if request("daytop1")=1 then
		daytop1=1
	else
		daytop1=0
	end if
	if request("weektop1")=1 then
		weektop1=1
	else
		weektop1=0
	end if
	if request("alltop1")=1 then
		alltop1=1
	else
		alltop1=0
	end if
	if request("allhot2")=1 then
		allhot2=1
	else
		allhot2=0
	end if
	if request("hot2")=1 then
		hot2=1
	else
		hot2=0
	end if
	if request("daytop2")=1 then
		daytop2=1
	else
		daytop2=0
	end if
	if request("weektop2")=1 then
		weektop2=1
	else
		weektop2=0
	end if
	if request("alltop2")=1 then
		alltop2=1
	else
		alltop2=0
	end if
	if request("allhot3")=1 then
		allhot3=1
	else
		allhot3=0
	end if
	if request("hot3")=1 then
		hot3=1
	else
		hot3=0
	end if
	if request("daytop3")=1 then
		daytop3=1
	else
		daytop3=0
	end if
	if request("weektop3")=1 then
		weektop3=1
	else
		weektop3=0
	end if
	if request("alltop3")=1 then
		alltop3=1
	else
		alltop3=0
	end if
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select * from showpage"
	rs.open sql,conndown,1,3
	rs("showset")=allhot1&","&hot1&","&daytop1&","&weektop1&","&alltop1&","&allhot2&","&hot2&","&daytop2&","&weektop2&","&alltop2&","&allhot3&","&hot3&","&daytop3&","&weektop3&","&alltop3
	rs.update
	rs.close
	set rs=nothing
	response.write("修改成功！")
end sub

sub save()
	dim SQL_Statement
	SQL_Statement=Request("SQL_Statement")
	if trim(SQL_Statement)<>"" then
		On Error Resume Next 
		conndown.Execute(SQL_Statement)
		if err.number="0" then
			response.write "执行成功"
		else
			response.write "语句有问题，具体出错如下："
			response.write Err.Description
			err.clear
		end if
	end if
end sub

sub renfile()
	dim filename,filenameold,filenamesplit,tempsplit,rannum
	dim downpath,fsdir,fsdir1,fs,file1,file2
	set rs=server.createobject("ADODB.Recordset")
	sql="select * from notdown"
	rs.open sql,conndown,1,3
	downpath=rs("downpath")
	rs.close
	sql="select * from download where ldown=1"
	rs.open sql,conndown,1,3
	do while not rs.eof
		filenameold=rs("filename")
		if InStr(filenameold,"/")<>0 or filenameold="" then
			response.write("文件名"&filenameold&"不正确！请仔细阅读说明")
		else
			filenamesplit=split(filenameold,".")
			tempsplit=split(filenamesplit(0),"$")
			do
				randomize
				rannum=right("000000"&int(999999*rnd),6)
				filename=tempsplit(0)&"$"&rannum
				for i=1 to ubound(filenamesplit)
					filename=filename&"."&filenamesplit(i)
				next
			loop while filename=filenameold
			fsdir=Server.MapPath("index.asp")	
			fsdir1=replace(fsdir,"\index.asp","\")
			set fs=createobject("scripting.filesystemobject") 
			file1=fsdir1&"\"&downpath&"\"&filenameold
			file2=fsdir1&"\"&downpath&"\"&filename
			If not fs.fileexists(file1) then
				response.write fsDir&"原文件不存在,请修改软件"&filenameold&"的文件名！<br>"
			elseIf fs.fileexists(file2) then
				response.write "目标文件"&filename&"已经存在,请重新改名！<br>"
			else 
				fs.MoveFile file1,file2
				If not fs.fileexists(file1) and fs.fileexists(file2) then
					response.write "文件名修改<font color=red>成功！</font><br>"
					rs("filename")=filename
					rs.update
				Else
					response.write "文件名修改<font color=red>失败！</font><br>"
				End If
			end if
		end if
		rs.movenext
	loop
	response.write("所有文件的改名操作成功！<br>")
	rs.close
	set rs=nothing
end sub%>
