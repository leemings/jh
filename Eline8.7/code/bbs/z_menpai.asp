<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="z_menpaiconfig.asp" -->
<!--#include file="z_plus_check.asp"-->
<%
stats="论坛门派"
call nav()
call head_var(2,0,"","")
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>您没有进入论坛门派的权限，请先登录或者同管理员联系。"
	call dvbbs_error()
else 
	call main()
end if
call activeonline()
call footer()
'============================================================
sub main()
%>
<table cellspacing=1 align=center border="0" class=tableborder1>
<tr><th valign=middle align=center height=24 width=<%=Forum_body(12)%> colspan="3" >门派说明</tr>
<tr>	
	<td valign=middle class=tablebody2 width="5%"></td>
	<td valign=middle class=tablebody1 width="90%">
	<ui>
		<li><u>明门正派</u>可以随意加入和退出
		<li>只有满足您要加入的<u>社团门派</u>的条件才可以加入，加入时要花费一定的金钱，但是魅力，经验，威望等会增加；退出时要被扣除几倍的入门费，并且要减一定的魅力，经验和威望
		<li>加入<u>魔教</u>金钱会增加，但是会减一定的魅力，经验和威望；退出时反之
		<li>被掌门踢出门派的用户会自动加入<u>中立派</u> 
		<li>加入一个新的门派前，请您先退出原有的门派，<u>中立派</u>除外 
	</ui>	
	</td>
	<td valign=middle class=tablebody2 width="5%"></td>
</tr>
</table>
<br>
<table cellspacing=1 align=center class=tableborder1 width=<%=Forum_body(12)%>>
	<tr><td height=20 class=tablebody2 colspan=9 align=center><a href=z_menpaicompare.asp><B>查看各门派对比</b></a></td></tr>
  <tr>
    <th height=24>门派名称</th>
    <th height=24>掌门</th> 
    <th height=24>创建日期</th> 
    <th height=24>性质</th> 
    <th height=24>弟子数</th> 
    <th height=24>金钱</th> 
    <th height=24>总帖数</th> 
    <th height=24>总积分</th> 
    <th height=24>操作</th> 
  </tr>
  <%
'门派显示  开始  
	dim rs1
	sql="select * from [GroupName]"
	set rs=conn.execute(sql)
	if rs.eof and rs.bof  then
		response.write "<tr><td height=21 class=tablebody1 colspan=8><font color=red>本论坛暂时没有任何门派</font></td></tr>"
	else 
		dim Zangmen,TotalArticle,TotalWealth,TotalUsernum,TotalScore
		dim menpai_setting

		do while not rs.eof
			TotalArticle=0
			TotalWealth=0
			TotalUsernum=0		
			TotalScore=0
			menpai_setting=split(rs("menpai_setting"),"|")
			Zangmen=trim(rs("Zangmen"))
			if Zangmen="" or isnull(Zangmen) then
				Zangmen="<font color=#FF6666>- 竞选中 -</font>"
			elseif Zangmen="-弃帮跑路-" then
				Zangmen="<font color=#CC6666>弃帮跑路</font>"
			else
				Zangmen="<a href=dispuser.asp?name="&htmlencode(Zangmen)&" title=""查看该掌门人的资料"" target=_blank>"&Zangmen&"</a>"
			end if	

			'统计指定门派的文章数、金钱数 和 弟子数（包括掌门）
			set rs1=conn.execute("select sum(article),sum(userwealth),sum(userscore),count(*) from [user] where usergroup='"&rs("GroupName")&"'")
			if rs1(0)<>"" then TotalArticle=rs1(0)
			if rs1(1)<>"" then TotalWealth=rs1(1)
			if rs1(2)<>"" then TotalScore=rs1(2)
			if rs1(3)<>"" then TotalUsernum=rs1(3)
%>
<tr>
	<td valign=middle align=center height=21 class=tablebody1><a href="z_menpaiadd.asp?menu=1&groupname=<%=rs("groupname")%>" title="门派简介:<%=menpai_setting(7)%>"><%=rs("groupname")%></a></td>
	<td valign=middle align=center height=21 class=tablebody1><%=Zangmen%></td>
    <td valign=middle align=center height=21 class=tablebody1><%=rs("addtime")%></td>
    <td valign=middle align=center height=21 class=tablebody1><%=MenpaiAttri(menpai_setting(0))%></td>
    <td valign=middle align=center height=21 class=tablebody1><%=TotalUsernum%></td>
    <td valign=middle align=center height=21 class=tablebody1><%=TotalWealth%></td>
    <td valign=middle align=center height=21 class=tablebody1><%=TotalArticle%></td>
    <td valign=middle align=center height=21 class=tablebody1><%=TotalScore%></td>
    <td valign=middle align=center height=21 class=tablebody1><%if MenpaiState(rs("state"))<>"正常" then%><%=MenpaiState(rs("state"))%><%else%><a href="z_menpaiadd.asp?menu=1&groupname=<%=rs("groupname")%>">查看</a><%end if%></td>
</tr>
<%			rs.movenext
		loop
		rs.close
	end if	  
'门派显示  结束	

'开创新门派部分  代码开始
	dim new_menpai_setting
	new_menpai_setting=split(conn.execute("select new_menpai_setting from [menpai]")(0),",")

	'set rs=conn.execute("select userwealth,usercp,userep,userpower,userclass from [user] where username='"&membername&"'")	
 	
	dim wumenwupai			'收纳不属于如何门派弟子的门派
	set rs=conn.execute("select GroupName from [GroupName] where state=1")  
	if rs.eof and rs.bof then 
		wumenwupai="无门无派"
	else
		wumenwupai=rs(0)
	end if
	rs.close
	
	dim mymenpai,CanCreateMenpai
	CanCreateMenpai=true
	set rs=conn.execute("select GroupName from [GroupName] where Zangmen='"&membername&"'")
	if rs.eof and rs.bof then
		mymenpai=conn.execute("select UserGroup from [user] where UserID="&UserID)(0)
		if mymenpai=wumenwupai or mymenpai="无门无派" then
			mymenpai="<font color=navy>"&mymenpai&"</font>"
		else
			mymenpai="<font color=navy>"&mymenpai&"</font><font color=red>弟子</font>"
		end if	
	else
		mymenpai="<font color=navy>"&rs(0)&"</font><font color=red>掌门</font>"
	end if    		
%>
<tr><td height=21 class=tablebody2 colspan=9></td></tr>
<tr><th colspan=9 height=20>申请建立门派</th></tr>
<tr>
	<td align=center height=21 class=tabletitle2>项目</td>
    <td align=center height=21 class=tabletitle2>金钱</td>
	<td align=center height=21 class=tabletitle2>魅力</td>
	<td align=center height=21 class=tabletitle2>经验</td>
	<td align=center height=21 class=tabletitle2>威望</td>
	<td align=center height=21 class=tabletitle2>文章</td>
	<td align=center height=21 class=tablebody2>用户等级</td>
	<td align=center height=21 class=tablebody2 colspan=2>门派</td>
</tr>
<tr>
	<td align=center height=21 class=tablebody2>创派要求</td>
    <td align=center height=21 class=tablebody1><%=new_menpai_setting(0)%>+<font color=red><%=new_menpai_setting(6)%></font></td>
	<td align=center height=21 class=tablebody1><%=new_menpai_setting(1)%></td>
	<td align=center height=21 class=tablebody1><%=new_menpai_setting(2)%></td>
	<td align=center height=21 class=tablebody1><%=new_menpai_setting(3)%></td>
	<td align=center height=21 class=tablebody1><%=new_menpai_setting(4)%></td>
	<td align=center height=21 class=tablebody1><%=GetUserTitle(new_menpai_setting(5))%>(<%=new_menpai_setting(5)%>级)</td>
	<td align=center height=21 class=tablebody1 colspan=2><%=wumenwupai%></td>
</tr>
<tr>
	<td align=center height=21 class=tablebody2>您的条件</td>
    <td align=center height=21 class=tablebody1><%=mymoney%></td>
	<td align=center height=21 class=tablebody1><%=myusercp%></td>
	<td align=center height=21 class=tablebody1><%=myuserep%></td>
	<td align=center height=21 class=tablebody1><%=mypower%></td>
	<td align=center height=21 class=tablebody1><%=myArticle%></td>
	<td align=center height=21 class=tablebody1><%=myClass%>(<%=GetUserClass(myClass)%>级)</td>
	<td align=center height=21 class=tablebody1 colspan=2><%=mymenpai%></td>
</tr>
<tr>
	<td align=center height=21 class=tablebody2>是否满足</td>
    <td align=center height=21 class=tablebody1><%if int(mymoney)>=(int(new_menpai_setting(0))+int(new_menpai_setting(6))) then%><font color=blue>√</font><%else%><%CanCreateMenpai=false%><font color=red>×</font><%end if%></td>
	<td align=center height=21 class=tablebody1><%if int(myusercp)>=int(new_menpai_setting(1)) then%><font color=blue>√</font><%else%><%CanCreateMenpai=false%><font color=red>×</font><%end if%></td>
	<td align=center height=21 class=tablebody1><%if int(myuserep)>=int(new_menpai_setting(2)) then%><font color=blue>√</font><%else%><%CanCreateMenpai=false%><font color=red>×</font><%end if%></td>
	<td align=center height=21 class=tablebody1><%if int(mypower)>=int(new_menpai_setting(3)) then%><font color=blue>√</font><%else%><%CanCreateMenpai=false%><font color=red>×</font><%end if%></td>
	<td align=center height=21 class=tablebody1><%if int(myArticle)>=int(new_menpai_setting(4)) then%><font color=blue>√</font><%else%><%CanCreateMenpai=false%><font color=red>×</font><%end if%></td>
	<td align=center height=21 class=tablebody1><%if int(GetUserClass(myClass))>=int(new_menpai_setting(5)) then%><font color=blue>√</font><%else%><%CanCreateMenpai=false%><font color=red>×</font><%end if%></td>
	<td align=center height=21 class=tablebody1 colspan=2><%if instr(mymenpai,wumenwupai)<>0 or instr(mymenpai,"无门无派")<>0 then%><font color=blue>√</font><%else%><%CanCreateMenpai=false%><font color=red>×</font><%end if%></td>
</tr>

<tr><td align=center height=21 class=tablebody1 colspan=9>
	<%if CanCreateMenpai then%>
		<a href="z_menpaiadd.asp?menu=8"><font color=blue>创建新门派</font></a>	
	<%else%>
		<font color=red>您至少有一个条件不满足创派要求</font>	
	<%end if%>
	</td>
</tr>
<tr><td align=center height=21 class=tablebody2 colspan=9></td></tr>
</table>
<%
'开创新门派部分  代码结束 
end sub
%>