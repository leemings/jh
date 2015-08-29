<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%     
'=========================================================
' File: boardpermission.asp
' Version:5.0 Final
' Date:  2002-10-20
' Script Written by sunwin
'=========================================================
 dim orders

	stats="论坛用户组权限查询"
	if boardid=0 then
	call nav()
	call head_var(2,0,"","")
	else
	call nav()
	call head_var(1,BoardDepth,0,0)
	end if
	if Cint(GroupSetting(39))=0 then
		Errmsg=Errmsg+"<br>"+"<li>您没有浏览本论坛用户组权限的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
		founderr=true
	end if
	if master then founderr=false
	if founderr then
		call dvbbs_error()
	else
           if not isInteger(request("orders")) or request("orders")="" then
		orders=1
	    else
		orders=request("orders")
	   end if
		call permission()
		if founderr then call dvbbs_error()
		call activeonline()
	end if
	call footer()
	REM 显示版面信息---Headinfo

function groupnum()
dim tmprs
	sql="Select count(*) from usergroups "
set tmprs=conn.execute(sql) 
groupnum=tmprs(0) 
set tmprs=nothing 
if isnull(groupnum) then groupnum=0
end function 

	sub permission()
dim trs,ars,Groupname
%>
<table class=tableborder1 cellspacing=1 cellpadding=3  align=center>
<tr>
<th height="25" width=16% id=tabletitlelink><a href=?boardid=<%=request("boardid")%>&orders=1 >浏览权限<a></th>
<th height="25" width=16% id=tabletitlelink><a href=?boardid=<%=request("boardid")%>&orders=2 >发帖权限<a></th>
<th height="25" width=16% id=tabletitlelink><a href=?boardid=<%=request("boardid")%>&orders=3 >帖子管理权限<a></th>
<th height="25" width=16% id=tabletitlelink><a href=?boardid=<%=request("boardid")%>&orders=4 >其他权限<a></th>
<th height="25" width=16% id=tabletitlelink><a href=?boardid=<%=request("boardid")%>&orders=5 >管理权限<a></th>
<th height="25" width=16% id=tabletitlelink><a href=?boardid=<%=request("boardid")%>&orders=6 >短信权限<a></th>
</tr>
</table>
<%
       select case orders
	case 1
%><table class=tableborder1 cellspacing=1 cellpadding=3  align=center>
<tr  align=center>
<td colspan=5 class=tablebody1>浏览权限</td>
</tr>
<tr>
<td height="25" width=20% class=tablebody2>用户组名(<%=groupnum()%>)</td>
<td height="25" width=20% class=tablebody2>可以浏览论坛</td>
<td height="25" width=20% class=tablebody2>可以查看会员信息(包括其他会员的资料和会员列表)</td>
<td height="25" width=20% class=tablebody2>可以查看其他人发布的主题</td>
<td height="25" width=20% class=tablebody2>可以浏览精华帖子</td>
</tr>
<%
	case 2
%><table class=tableborder1 cellspacing=1 cellpadding=3  align=center>
<tr  align=center>
<td colspan=14 class=tablebody1>发帖权限</td>
</tr>
<tr>
<td height="25" width=22% class=tablebody2>用户组名(<%=groupnum()%>)</td>
<td height="25" width=6% class=tablebody2>可以发布新主题</td>
<td height="25" width=6% class=tablebody2>可以回复自己的主题</td>
<td height="25" width=6% class=tablebody2>可以回复其他人的主题</td>
<td height="25" width=6% class=tablebody2>可以在论坛允许评分的时候参与评分(鲜花和鸡蛋)?</td>
<td height="25" width=6% class=tablebody2>参与评分所需金钱</td>
<td height="25" width=6% class=tablebody2>可以上传附件</td>
<td height="25" width=6% class=tablebody2>一次最多上传文件个数</td>
<td height="25" width=6% class=tablebody2>一天最多上传文件个数</td>
<td height="25" width=6% class=tablebody2>上传文件大小限制</td>
<td height="25" width=6% class=tablebody2>可以发布新投票</td>
<td height="25" width=6% class=tablebody2>可以参与投票</td>
<td height="25" width=6% class=tablebody2>可以发布小字报</td>
<td height="25" width=6% class=tablebody2>发布小字报所需金钱</td>
</tr>
<%
	case 3
%><table class=tableborder1 cellspacing=1 cellpadding=3  align=center>
<tr  align=center>
<td colspan=5 class=tablebody1>帖子管理权限</td>
</tr>
<tr>
<td height="25" width=20% class=tablebody2>用户组名(<%=groupnum()%>)</td>
<td height="25" width=20% class=tablebody2>可以编辑自己的帖子</td>
<td height="25" width=20% class=tablebody2>可以删除自己的帖子</td>
<td height="25" width=20% class=tablebody2>可以移动自己的帖子到其他论坛</td>
<td height="25" width=20% class=tablebody2>可以打开/关闭自己发布的主题</td>
</tr>
<%
	case 4
%><table class=tableborder1 cellspacing=1 cellpadding=3  align=center>
<tr  align=center>
<td colspan=11 class=tablebody1>其他权限</td>
</tr>
<tr>
<td height="25" width=* class=tablebody2>用户组名(<%=groupnum()%>)</td>
<td height="25" width=9% class=tablebody2>可以搜索论坛</td>
<td height="25" width=9% class=tablebody2>可以使用'发送本页给好友'功能</td>
<td height="25" width=9% class=tablebody2>可以修改个人资料</td>
<td height="25" width=9% class=tablebody2>是否有审核帖子的权限</td>
<td height="25" width=9% class=tablebody2>是否有进入隐含论坛的权限</td>
<td height="25" width=9% class=tablebody2>是否有论坛文件管理权限</td>
<td height="25" width=9% class=tablebody2>可以进行帖子总固顶操作</td>
<td height="25" width=9% class=tablebody2>可以浏览论坛事件</td>
<td height="25" width=9% class=tablebody2>可以浏览论坛展区</td>
<td height="25" width=9% class=tablebody2>是否有帖子主题颜色权限</td>
</tr>
<%
	case 5
%><table class=tableborder1 cellspacing=1 cellpadding=3  align=center>
<tr  align=center>
<td colspan=18 class=tablebody1>管理权限</td>
</tr>
<tr>
<td height="25" width=10% class=tablebody2>用户组名(<%=groupnum()%>)</td>
<td height="25" width=5% class=tablebody2>可以删除其它人帖子</td>
<td height="25" width=5% class=tablebody2>可以移动其它人帖子</td>
<td height="25" width=5% class=tablebody2>可以打开/关闭其它人帖子</td>
<td height="25" width=5% class=tablebody2>可以固顶/解除固顶帖子</td>
<td height="25" width=5% class=tablebody2>可以奖励/惩罚发帖用户</td>
<td height="25" width=5% class=tablebody2>可以奖励/惩罚用户</td>
<td height="25" width=5% class=tablebody2>可以编辑其它人帖子</td>
<td height="25" width=5% class=tablebody2>可以加入/解除精华帖子</td>
<td height="25" width=5% class=tablebody2>可以发布公告</td>
<td height="25" width=5% class=tablebody2>可以管理公告</td>
<td height="25" width=5% class=tablebody2>可以管理小字报</td>
<td height="25" width=5% class=tablebody2>可以锁定/屏蔽/解除锁定用户</td>
<td height="25" width=5% class=tablebody2>可以删除用户1－10天内所发帖子</td>
<td height="25" width=5% class=tablebody2>可以查看来访IP及来源</td>
<td height="25" width=5% class=tablebody2>可以限定IP来访</td>
<td height="25" width=5% class=tablebody2>可以管理用户权限</td>
<td height="25" width=5% class=tablebody2>可以批量删除帖子（前台）</td>
</tr>
<%
	case 6
%><table class=tableborder1 cellspacing=1 cellpadding=3  align=center>
<tr  align=center>
<td colspan=5 class=tablebody1>短信权限</td>
</tr>
<tr>
<td height="25" width=12.5% class=tablebody2>用户组名(<%=groupnum()%>)</td>
<td height="25" width=12.5% class=tablebody2>可以发送短信</td>
<td height="25" width=12.5% class=tablebody2>最多发送用户</td>
<td height="25" width=12.5% class=tablebody2>短信内容大小限制</td>
<td height="25" width=12.5% class=tablebody2>信箱大小限制</td>
</tr>
<%
	case else
%><table class=tableborder1 cellspacing=1 cellpadding=3  align=center>
<tr  align=center>
<td colspan=5 class=tablebody1>其他权限</td>
</tr>
<tr>
<td height="25" width=20% class=tablebody2>用户组名(<%=groupnum()%>)</td>
<td height="25" width=20% class=tablebody2>可以浏览论坛</td>
<td height="25" width=20% class=tablebody2>可以查看会员信息(包括其他会员的资料和会员列表)</td>
<td height="25" width=20% class=tablebody2>可以查看其他人发布的主题</td>
<td height="25" width=20% class=tablebody2>可以浏览精华帖子</td>
</tr>
<%	end select

	set trs=conn.execute("select * from usergroups order by usergroupid")
        do while not trs.eof
        groupname=trs("title")

	set ars=conn.execute("select * from BoardPermission where BoardID="&boardid&" and GroupID="&trs("UserGroupID"))
          if  not ars.eof then
        GroupSetting=split(ars("PSetting"),",")
          else
         GroupSetting=split(trs("GroupSetting"),",")
         end if

       select case orders
	case 1
%>
<tr> 
<td height="25"  class=tablebody1><%=groupname%></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(0)=1 then%>√<%end if%><%if GroupSetting(0)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(1)=1 then%>√<%end if%><%if GroupSetting(1)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(2)=1 then%>√<%end if%><%if GroupSetting(2)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(41)=1 then%>√<%end if%><%if GroupSetting(41)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<%
	case 2
%>
<tr>
<td height="25"  class=tablebody1><%=groupname%></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(3)=1 then%>√<%end if%><%if GroupSetting(3)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(4)=1 then%>√<%end if%><%if GroupSetting(4)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(5)=1 then%>√<%end if%><%if GroupSetting(5)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(6)=1 then%>√<%end if%><%if GroupSetting(6)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%=GroupSetting(47)%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(7)=1 then%>√<%end if%><%if GroupSetting(7)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%><%if GroupSetting(7)=2 then%><font color=<%=Forum_body(8)%>>发</font><%end if%><%if GroupSetting(7)=3 then%><font color=<%=Forum_body(8)%>>回</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%=GroupSetting(40)%></B></td>
<td height="25"  class=tablebody1><B><%=GroupSetting(50)%></B></td>
<td height="25"  class=tablebody1><B><%=GroupSetting(44)%></B> KB</td>
<td height="25"  class=tablebody1><B><%if GroupSetting(8)=1 then%>√<%end if%><%if GroupSetting(8)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(9)=1 then%>√<%end if%><%if GroupSetting(9)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(17)=1 then%>√<%end if%><%if GroupSetting(17)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%=GroupSetting(46)%></B></td>
</tr>
<%
	case 3
%>
<tr>
<td height="25"  class=tablebody1><%=groupname%></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(10)=1 then%>√<%end if%><%if GroupSetting(10)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(11)=1 then%>√<%end if%><%if GroupSetting(11)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(12)=1 then%>√<%end if%><%if GroupSetting(12)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(13)=1 then%>√<%end if%><%if GroupSetting(13)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<%
	case 4
%>
<tr>
<td height="25"  class=tablebody1><%=groupname%></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(14)=1 then%>√<%end if%><%if GroupSetting(14)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(15)=1 then%>√<%end if%><%if GroupSetting(15)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(16)=1 then%>√<%end if%><%if GroupSetting(16)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(36)=1 then%>√<%end if%><%if GroupSetting(36)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(37)=1 then%>√<%end if%><%if GroupSetting(37)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(48)=1 then%>√<%end if%><%if GroupSetting(48)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(38)=1 then%>√<%end if%><%if GroupSetting(38)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(39)=1 then%>√<%end if%><%if GroupSetting(39)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(49)=1 then%>√<%end if%><%if GroupSetting(49)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(51)=1 then%>√<%end if%><%if GroupSetting(51)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<%
	case 5
%>
<tr>
<td height="25"  class=tablebody1><%=groupname%></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(18)=1 then%>√<%end if%><%if GroupSetting(18)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(19)=1 then%>√<%end if%><%if GroupSetting(19)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(20)=1 then%>√<%end if%><%if GroupSetting(20)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(21)=1 then%>√<%end if%><%if GroupSetting(21)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(22)=1 then%>√<%end if%><%if GroupSetting(22)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(43)=1 then%>√<%end if%><%if GroupSetting(43)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(23)=1 then%>√<%end if%><%if GroupSetting(23)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(24)=1 then%>√<%end if%><%if GroupSetting(24)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(25)=1 then%>√<%end if%><%if GroupSetting(25)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(26)=1 then%>√<%end if%><%if GroupSetting(26)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(27)=1 then%>√<%end if%><%if GroupSetting(27)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(28)=1 then%>√<%end if%><%if GroupSetting(28)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(29)=1 then%>√<%end if%><%if GroupSetting(29)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(30)=1 then%>√<%end if%><%if GroupSetting(30)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(31)=1 then%>√<%end if%><%if GroupSetting(31)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(42)=1 then%>√<%end if%><%if GroupSetting(42)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(45)=1 then%>√<%end if%><%if GroupSetting(45)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<%
	case 6
%>
<tr>
<td height="25"  class=tablebody1><%=groupname%></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(32)=1 then%>√<%end if%><%if GroupSetting(32)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%=GroupSetting(33)%></B></td>
<td height="25"  class=tablebody1><B><%=GroupSetting(34)%></B> byte</td>
<td height="25"  class=tablebody1><B><%=GroupSetting(35)%></B> KB</td>
</tr>
<%
	case else
%>
<tr> 
<td height="25"  class=tablebody1><%=groupname%></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(0)=1 then%>√<%end if%><%if GroupSetting(0)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(1)=1 then%>√<%end if%><%if GroupSetting(1)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(2)=1 then%>√<%end if%><%if GroupSetting(2)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(41)=1 then%>√<%end if%><%if GroupSetting(41)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<%
	end select
trs.movenext
loop
set trs=nothing
set ars=nothing
%>
</table>
<% end sub %>