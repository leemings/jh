<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="md5.asp"-->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<script language="JavaScript">
<!--
function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name != 'chkall')
       e.checked = form.chkall.checked;
    }
  }
//-->
</script>
<%
	dim admin_flag
	admin_flag="26"
	if not master or instr(session("flag"),admin_flag)=0 then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		dim body,username2,password2,oldpassword,oldusername,oldadduser,username1
'''''''''''''''
'取出用户组管理员的组名 2002-12-13
		dim groupsname,titlepic
		set rs=conn.execute("select title from [UserGroups] where UserGroupID=1 ")
		groupsname=rs(0)
		set rs=conn.execute("select titlepic from usertitle where usergroupid=1 order by Minarticle desc")
		titlepic=rs(0)
		set rs=nothing

		if request("action")="updat" then
			call update()
			response.write body
		elseif request("action")="del" then
			call del()
			response.write body
        	elseif request("action")="pasword" then
			call pasword()
        	elseif request("action")="newpass" then
			call newpass()
			response.write body
		elseif request("action")="add" then
			call addadmin()
		elseif request("action")="edit" then
			call userinfo()
		elseif request("action")="savenew" then
			call savenew()
			response.write body
		else
			call userlist()
		end if
		conn.close
		set conn=nothing
	end if

	sub userlist()
%>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
                <tr> 
                  <th height=22 colspan=5>管理员管理(点击用户名进行操作)</th>
                </tr>
                <tr align=center> 
                  <td width="30%" height=22 class="forumHeaderBackgroundAlternate"><B>用户名</B></td><td width="25%" class="forumHeaderBackgroundAlternate"><B>上次登录时间</B></td><td width="15%" class="forumHeaderBackgroundAlternate"><B>上次登录IP</B></td><td width="15%" class="forumHeaderBackgroundAlternate"><B>操作</B></td>
                </tr>
<%
	set rs=conn.execute("select * from admin order by LastLogin desc")
	do while not rs.eof
%>
                <tr> 
                  <td class=forumrow><a href="admin_admin.asp?id=<%=rs("id")%>&action=pasword"><%=rs("username")%></a></td><td class=forumrow><%=rs("LastLogin")%></td><td class=forumrow><%=rs("LastLoginIP")%></td><td class=forumrow><a href="admin_admin.asp?action=del&id=<%=rs("id")%>&name=<%=rs("username")%>">删除</a>&nbsp;&nbsp;<a href="admin_admin.asp?id=<%=rs("id")%>&action=edit">编辑权限</a></td>
                </tr>
<%
	rs.movenext
	loop
	rs.close
	set rs=nothing
%>
	       </table>
<%
	end sub

	sub del()
	conn.execute("delete from admin where id="&request("id"))
	conn.execute("update [user] set usergroupid=4 where username='"&replace(request("name"),"'","")&"'")
	body="<li>管理员删除成功。"
	end sub

sub pasword()
	set rs=conn.execute("select * from admin where id="&request("id"))
	oldpassword=rs("password")
	oldadduser=rs("adduser")
  %> 
<form action="?action=newpass" method=post>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
               <tr> 
                  <th colspan=2 height=23>管理员资料管理－－密码修改
                  </th>
                </tr>
               <tr > 
            <td width="26%" align="right" class=forumrow>后台登录名称：</td>
            <td width="74%" class=forumrow>
              <input type=hidden name="oldusername" value="<%=rs("username")%>">
              <input type=text name="username2" value="<%=rs("username")%>">  (可与注册名不同)
            </td>
          </tr>
          <tr > 
            <td width="26%" align="right" class=forumrow>后台登录密码：</td>
            <td width="74%" class=forumrow>
              <input type="password" name="password2" value="<%=oldpassword%>">  (可与注册密码不同,如要修改请直接输入)
            </td>
          </tr>
          <tr > 
            <td width="26%" align="right" class=forumrow height=23>前台用户名称：</td>
            <td width="74%" class=forumrow><%=oldadduser%>
            </td>
          </tr>
          <tr align="center"> 
            <td colspan="2" class=forumrow> 

              <input type=hidden name="oldadduser" value="<%=oldadduser%>">
              <input type=hidden name="adduser" value="<%=oldadduser%>">
              <input type=hidden name=id value="<%=request("id")%>">
              <input type="submit" name="Submit" value="更 新">
            </td>
          </tr>
        </table>
        </form>

<%       rs.close
         set rs=nothing
end sub

sub newpass()
	dim passnw,usernw,aduser
	set rs=conn.execute("select * from admin where id="&request("id"))
	oldpassword=rs("password")
	if request("username2")="" then
		body="<li>请输入管理员名字。<a href=?>［ <font color=red>返回</font> ］</a>"
		exit sub
	else 
		usernw=trim(request("username2"))
	end if
	if request("password2")="" then
		body="<li>请输入您的密码。<a href=?>［ <font color=red>返回</font> ］</a>"
		exit sub
	elseif trim(request("password2"))=oldpassword then
		passnw=request("password2")
	else
		passnw=md5(request("password2"))
	end if
	if request("adduser")="" then
		body="<li>请输入管理员名字。<a href=?>［ <font color=red>返回</font> ］</a>"
		exit sub
	else 
		aduser=trim(request("adduser"))
	end if

	set rs=server.createobject("adodb.recordset")
	sql="select * from admin where username='"&trim(request("oldusername"))&"'"
	rs.open sql,conn,1,3
	if not rs.eof and not rs.bof then
	rs("username")=usernw
	rs("adduser")=aduser
	rs("password")=passnw
''''''''''''''
'更新用户的的级别
        conn.execute("update [user] set usergroupid=1,userclass='"&groupsname&"',titlepic='"&titlepic&"' where username='"&trim(request("oldusername"))&"'")	'
	body="<li>管理员资料更新成功，请记住更新信息。<br> 管理员："&request("username2")&" <BR> 密   码："&request("password2")&" <a href=?>［ <font color=red>返回</font> ］</a>"
	rs.update
	
	end if
	rs.close
	set rs=nothing
end sub


sub addadmin()
%> 
<form action="?action=savenew" method=post>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
               <tr> 
                  <th colspan=2 height=23>管理员管理－－添加管理员
                  </th>
                </tr>
               <tr > 
            <td width="26%" align="right" class=forumrow>后台登录名称：</td>
            <td width="74%" class=forumrow>
              <input type=text name="username2">  (可与注册名不同)
            </td>
          </tr>
          <tr > 
            <td width="26%" align="right" class=forumrow>后台登录密码：</td>
            <td width="74%" class=forumrow>
              <input type="password" name="password2">  (可与注册密码不同)
            </td>
          </tr>
          <tr > 
            <td width="26%" align="right" class=forumrow height=23>前台用户名称：</td>
            <td width="74%" class=forumrow><input type=text name="username1">  (本选项填写后不允许修改)
            </td>
          </tr>
          <tr align="center"> 
            <td colspan="2" class=forumrow> 
              <input type="submit" name="Submit" value="添 加">
            </td>
          </tr>
        </table>
        </form>

<%
end sub

sub savenew()
dim adminuserid
	if request.form("username2")="" then
	body="请输入后台登录用户名！"
	exit sub
	end if
	if request.form("username1")="" then
	body="请输入前台登录用户名！"
	exit sub
	end if
	if request.form("password2")="" then
	body="请输入后台登录密码！"
	exit sub
	end if

	set rs=conn.execute("select userid from [user] where username='"&replace(request.form("username1"),"'","")&"'")
	if rs.eof and rs.bof then
	body="您输入的用户名不是一个有效的注册用户！"
	exit sub
        else
        adminuserid=rs(0)
	end if

	set rs=conn.execute("select username from admin where username='"&replace(request.form("username2"),"'","")&"'")
	if not (rs.eof and rs.bof) then
	body="您输入的用户名已经在管理用户中存在！"
	exit sub
	end if
	conn.execute("update [user] set usergroupid=1 , userclass='"&groupsname&"',titlepic='"&titlepic&"' where userid="&adminuserid&" ")
	conn.execute("insert into admin (username,[password],adduser) values ('"&replace(request.form("username2"),"'","")&"','"&md5(replace(request.form("password2"),"'",""))&"','"&replace(request.form("username1"),"'","")&"')")
	body="用户ID:"&adminuserid&" 添加成功，请记住新管理员后台登录信息，如需修改请返回管理员管理！"
end sub

sub userinfo()
dim menu(7,10)
menu(0,0)="常规设置"
menu(0,1)="<a href=admin_setting.asp target=main>常规设置信息</a>"
menu(0,2)="<a href=admin_ads.asp target=main>论坛广告设置</a>"
menu(0,3)="<a href=bbseven.asp?action=batch target=_blank>论坛日志管理</a>"
menu(0,4)="<a href=admin_inputfile.asp target=main>初始信息设置</a>"

menu(1,0)="论坛管理"
menu(1,1)="<a href=admin_board.asp?action=add target=main>论坛版面添加</a> | <a href=admin_board.asp target=main>管理</a>"
menu(1,2)="<a href=admin_board.asp?action=permission target=main>论坛权限管理</a>"
menu(1,3)="<a href=admin_boardunite.asp target=main>合并论坛数据</a>"
menu(1,4)="<a href=admin_update.asp target=main>更新论坛数据</a>"
menu(1,5)="<a href=z_link_post.asp target=_blank>联盟论坛管理</a>"
menu(1,6)="<a href=announcements.asp?boardid=0&action=AddAnn target=_blank>发布论坛公告 | 管理</a>"
menu(1,7)="<a href=z_admin_menpai.asp target=main>论坛门派管理</a>"

menu(2,0)="用户管理"
menu(2,1)="<a href=admin_user.asp target=main>用户信息管理</a>"
menu(2,2)="<a href=admin_grade.asp?action=add target=main>论坛等级添加</a> | <a href=admin_grade.asp target=main>管理</a>"
menu(2,3)="<a href=admin_wealth.asp target=main>用户积分设置</a>"
menu(2,4)="<a href=admin_message.asp target=main>论坛短信管理</a>"
menu(2,5)="<a href=admin_group.asp?action=addgroup target=main>用户组添加</a> | <a href=admin_group.asp target=main>管理</a>"
menu(2,6)="<a href=admin_admin.asp?action=add target=main>管理员添加</a> | <a href=admin_admin.asp target=main>管理</a>"
menu(2,7)="<a href=admin_mailist.asp target=main>邮件列表</a> | <a href=admin_mailout.asp target=main>列表导出</a>"
menu(2,8)="<a href=admin_update.asp?action=updateuser target=main>更新用户数据</a>"

menu(3,0)="主题和帖子设置"
menu(3,1)="<a href=admin_alldel.asp target=main>批量删除</a>"
menu(3,2)="<a href=admin_alldel.asp?action=moveinfo target=main>批量移动</a>"
menu(3,3)="<a href=recycle.asp target=_blank>回收站管理</a>"
menu(3,4)="<a href=admin_postdata.asp?action=Nowused target=main>当前帖子数据表管理</a>"
menu(3,5)="<a href=admin_postdata.asp target=main>数据表间帖子转换</a>"

menu(4,0)="外观设置"
menu(4,1)="<a href=admin_color.asp target=main>论坛风格CSS设置</a>"
menu(4,2)="<a href=admin_pic.asp target=main>基本图片设置</a>"
menu(4,3)="<a href=admin_skin.asp?action=news target=main>设置模板添加</a> | <a href=admin_skin.asp target=main>管理</a>"
menu(4,4)="<a href=admin_loadskin.asp target=main>CSS风格导出</a> | <a href=admin_loadskin.asp?action=load target=main>导入</a>"

menu(5,0)="替换/限制处理"
menu(5,1)="<a href=admin_badword.asp?reaction=badword target=main>帖子过滤字符</a>"
menu(5,2)="<a href=admin_badword.asp?reaction=splitreg target=main>注册过滤字符</a>"
menu(5,3)="<a href=admin_lockip.asp?action=add target=main>IP来访限定添加</a> | <a href=admin_lockip.asp target=main>管理</a>"
menu(5,4)="<a href=admin_address.asp?action=add target=main>论坛IP库添加</a> | <a href=admin_address.asp target=main>管理</a>"
menu(5,5)="<a href=admin_address.asp?action=upip target=main>导入IP库</a>"

menu(6,0)="数据处理(Access)"
menu(6,1)="<a href=admin_data.asp?action=CompressData target=main>压缩数据库</a>"
menu(6,2)="<a href=admin_data.asp?action=BackupData target=main>备份数据库</a>"
menu(6,3)="<a href=admin_data.asp?action=RestoreData target=main>恢复数据库</a>"
menu(6,4)="<a href=admin_data.asp?action=SpaceSize target=main>系统空间占用</a>"

menu(7,0)="文件管理"
menu(7,1)="<a href=admin_upUserface.asp target=main>上传头像管理</a>"
menu(7,2)="<a href=admin_uploadlist.asp target=main>上传文件管理</a>"
menu(7,3)="<a href=admin_bbsface.asp?orders=1 target=main>注册头像管理</a>"
menu(7,4)="<a href=admin_bbsface.asp?orders=2 target=main>发帖心情管理</a>"
menu(7,5)="<a href=admin_bbsface.asp?orders=3 target=main>发帖表情管理</a>"
dim j,tmpmenu,menuname,menurl
set rs=conn.execute("select * from admin where id="&request("id"))
%>
<form action="admin_admin.asp?action=updat" method=post name=adminflag>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
<tr> 
<th height=25><b>管理员权限管理</b>(请选择相应的权限分配给管理员<%=rs("username")%>)
</th>
</tr>
<tr> 
<td height=25 class="forumHeaderBackgroundAlternate"><b>>>全局权限</b></td></tr>
<tr><td class=forumrow>
<%for i=0 to ubound(menu,1)%>
<b><%=menu(i,0)%></b><br>
<%
on error resume next
for j=1 to ubound(menu,2)
if isempty(menu(i,j)) then exit for
tmpmenu=split(menu(i,j),",")
menuname=tmpmenu(0)
menurl=tmpmenu(1)
%>
<input type="checkbox" name="flag" <% if instr(session("flag"),"26")=0 then response.write "disabled=true" %> value="<%=i&j%>" <% if instr(rs("flag"),i&j)<>0 then response.write "checked" %>><a href="<%=menurl%>"><%=menuname%></a>&nbsp;&nbsp;
<%next%><br><br>
<%next%>
<input type=hidden name=id value="<%=request("id")%>">
<input type="submit" name="Submit" value="更新"><input name=chkall type=checkbox value=on onclick=CheckAll(this.form)>选择所有权限
</td>
</tr>
</table>
</form>
<%
rs.close
set rs=nothing
end sub

sub update()
'00,01,02,03,10,11,12,13,14,15,16,17,20,21,22,23,24,25,26,27,28,30,31,32,33,34,35,40,41,42,43,44,50,51,52,53,54,55,60,61,62,63,64,70,71,72,73,74,75
set rs=server.createobject("adodb.recordset")
sql="select * from admin where id="&request("id")
rs.open sql,conn,1,3
if not rs.eof and not rs.bof then
rs("flag")=request("flag")
body="<li>管理员更新成功，请记住更新信息。"
rs.update
if rs("username")=membername then session("flag")=request("flag")
end if
rs.close
set rs=nothing
end sub
%>