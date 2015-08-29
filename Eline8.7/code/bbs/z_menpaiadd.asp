<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="z_menpaiconfig.asp" -->
<%
'========================主程序开始===========================
	dim menu,GroupName,chenbackasp,chentitle
	menu=request.querystring("menu")
	GroupName=checkstr(Trim(Request("GroupName")))
	if menu="" then menu=0
	
	chenbackasp="z_menpai.asp"
	chentitle="论坛门派"
	stats="论坛门派 "&htmlencode(GroupName)	
	call nav()
	
	select case menu
		case 2
			stats="论坛门派 加入门派 "&htmlencode(GroupName)
		case 3
			stats="论坛门派 退出门派 "&htmlencode(GroupName)
		case 4,5,6,9
			chenbackasp="z_menpaiadd.asp?GroupName="&htmlencode(GroupName)
			chentitle=htmlencode(GroupName)
			stats=htmlencode(GroupName)&" 门派管理"						
		case 7,8
			stats="论坛门派 申请新门派"
	end select
	call head_var(3,0,"","")
			
	if not founduser then
		stats="论坛门派 错误信息"
		Errmsg=Errmsg+"<br>"+"<li>您没有进入论坛门派的权限，请先登录或者同管理员联系。"
		call dvbbs_error()
	else
		'读取指定门派的值
		dim Menpai_setting,Zangmen,ZangmenStr,Addtime,States 
		dim CanManage		'能否管理本派，必须本派掌门才能管理本派
		CanManage=false 
		
		if GroupName<>"" and (not isnull(GroupName)) then
			sql="select * from [GroupName] where groupname='"&GroupName&"'"
			set rs=conn.execute(sql)
			if not(rs.eof and rs.bof) then
				GroupName=rs("GroupName")
				Zangmen=trim(rs("Zangmen"))
				Addtime=rs("Addtime")
				States=rs("State")
				Menpai_setting=split(rs("Menpai_setting"),"|")
				
				if Zangmen="" or isnull(Zangmen) then
					ZangmenStr="<font color=#FF6666>- 竞选中 -</font>"
				elseif Zangmen="-弃帮跑路-" then
					ZangmenStr="<font color=#CC6666>弃帮跑路</font>"
				else
					if Zangmen=membername then CanManage=true
					ZangmenStr="<a href=dispuser.asp?name="&htmlencode(Zangmen)&" title=""查看该掌门人的资料"" target=_blank>"&Zangmen&"</a>"
				end if				
			else
				GroupName=""
				Menpai_setting=split("-1|0|0|0|0|0|1|","|")
			end if	
			rs.close
		else
			GroupName=""
			Menpai_setting=split("-1|0|0|0|0|0|1|","|")
		end if
		
		dim wumenwupai			'收纳不属于如何门派弟子的门派
		set rs=conn.execute("select GroupName from [GroupName] where state=1")  
		if rs.eof and rs.bof then 
			wumenwupai="无门无派"
		else
			wumenwupai=rs(0)
		end if
		rs.close
		
		dim mymenpai		'自己所在的门派	
		mymenpai=conn.execute("select UserGroup from [user] where username='"&membername&"'")(0)				 	
		
		select case menu
			case 1
				call menpai()		'显示门派的基本情况
			case 2
				call inmenpai()		'加入门派
			case 3  	
    			call exitmenpai()	'退出门派 
    		case 4  						
				call manage() 		'管理门派 － 修改门派资料页面 
			case 5 
		    	call clearmember()		'清理门户  
			case 6  
			   	call changegroup()	    '保存修改（门派管理）
			case 7  
				call applynewmenpai()	'申请新门派－写入数据库
			case 8
    			call newmenpai()		'申请新门派－界面 
			case 9
    			call logout()			'注销门派	 				
			case else
				call menpai()
		end select	
	end if
	if founderr then call menpai_err()
	call activeonline()
	call footer() 
'========================主程序结束===========================

'---------------------------显示门派资料----------------------
sub menpai()
	if GroupName="" then 
		Errmsg=Errmsg+"<br>"+"<li>错误的门派参数，请指定相关门派"
		founderr=true
		exit sub
	end if 
	dim TotalArticle,TotalWealth,TotalUsernum,CanJoinin,CanExitout
	TotalArticle=0:	TotalWealth=0:	TotalUsernum=0
	CanJoinin=true:		CanExitout=false
	'统计指定门派的文章数、金钱数 和 弟子数（包括掌门） 
	set rs=conn.execute("select sum(article),sum(userwealth),count(*) from [user] where usergroup='"&GroupName&"'") 
	if rs(0)<>"" then TotalArticle=rs(0)
	if rs(1)<>"" then TotalWealth=rs(1)
	if rs(2)<>"" then TotalUsernum=rs(2)
	
	dim yourmsg
%>
<table cellspacing=1 align=center class=tableborder1 width=<%=Forum_body(12)%>>
	<tr>
		<th valign=middle align=center height=24 colspan=8 id=TableTitleLink>门派资料 | <a href=z_menpaimember.asp?Groupname=<%=Groupname%>>查看本派弟子</a></th> 
	</tr> 
	<tr> 
		<td valign=middle align=center class=tabletitle2 height="22">门派名称</td> 
		<td valign=middle align=center class=tabletitle2 height="22">掌门</td> 
		<td valign=middle align=center class=tabletitle2 height="22">性质</td> 
		<td valign=middle align=center class=tabletitle2 height="22">创建日期</td>
		<td valign=middle align=center class=tabletitle2 height="22">弟子数</td>
		<td valign=middle align=center class=tabletitle2 height="22">金钱数</td>
		<td valign=middle align=center class=tabletitle2 height="22">文章数</td>
		<td valign=middle align=center class=tabletitle2 height="22">当前状态</td>
	</tr> 
	<tr> 
		<td valign=middle align=center class=tablebody1 height="20"><font color="#FF9933"><font face=Wingdings color=blue>[</font> <%=GroupName%></font></td>   	
		<td valign=middle align=center class=tablebody1 height="20"><%=ZangmenStr%></td>
		<td valign=middle align=center class=tablebody1 height="20"><%=MenpaiAttri(Menpai_setting(0))%></td> 
		<td valign=middle align=center class=tablebody1 height="20"><%=Addtime%></td>
		<td valign=middle align=center class=tablebody1 height="20"><%=TotalUsernum%></td> 
		<td valign=middle align=center class=tablebody1 height="20"><%=TotalWealth%></td> 
		<td valign=middle align=center class=tablebody1 height="20"><%=TotalArticle%></td> 
		<td valign=middle align=center class=tablebody1 height="20"><%=MenpaiState(States)%></td>
	</tr> 
	<tr><td height=8 class=tablebody2 colspan=8></td></tr>
	<tr> 
		<td valign=middle align=center class=tabletitle2 height="22">门派简介</td>  
		<td valign=middle align=left class=tablebody1 height="22" colspan=7>&nbsp;<font face=Wingdings color=blue>v</font>&nbsp;<%=menpai_setting(7)%></td> 
	</tr> 
	<tr><td height=8 class=tablebody2 colspan=8></td></tr>
	<tr> 
		<td valign=middle align=center class=tabletitle2 height="22">门规</td>  
		<td valign=middle align=left class=tablebody1 height="22" colspan=7><ui>
<%		select case Menpai_setting(0)
			case 0
				response.Write "<li>本派是<u>明门正派</u>，您可以随意加入和退出本派"
			case 1
				response.Write "<li>本派是<u>社团门派</u>，必须满足条件才能加入和退出本派"
				response.Write "<li>加入本派要花费一定的金钱，但是魅力，经验，威望等会增加"
				response.Write "<li>退出本派要扣除几倍的入门费，并且要减一定的魅力，经验和威望"			
			case 2
				response.Write "<li>本派是<u>魔教</u>(在那些所谓的明门正派眼里)，必须满足条件才能加入和退出本派"
				response.Write "<li>加入本派您的金钱会增加，但是会减一定的魅力，经验和威望；退出时反之"
			case 3
				response.Write "<li><u>中立派</u>，不属于任何门派，群蛇混杂，什么人都有"
		end select 
		if cint(Menpai_setting(0))<>3 then
			response.Write "<li>掌门有权将不遵守门规的人踢出本派，被踢出门派的用户会自动加入<u>中立派</u>"
		end if	
%>
		</ui></td> 
	</tr> 		
	<tr><td height=15 class=tablebody2 colspan=8></td></tr>
	<th height="22" colspan=8>加入本门派条件</th>
	<tr>
		<td align=center height=21 width="15%" class=tabletitle2>项目</td> 
		<td align=center height=21 class=tabletitle2>金钱</td> 
		<td align=center height=21 class=tabletitle2>经验</td>
		<td align=center height=21 class=tabletitle2>魅力</td> 
		<td align=center height=21 class=tabletitle2>威望</td>
		<td align=center height=21 class=tabletitle2>文章</td> 
		<td align=center height=21 class=tablebody2>用户等级</td>
		<td align=center height=21 class=tablebody2>门派</td>
	</tr>	
	<tr>
		<td align=center height=21 class=tablebody2>入派要求</td>
		<td align=center height=21 class=tablebody1><%=menpai_setting(1)%></td>
		<td align=center height=21 class=tablebody1><%=menpai_setting(2)%></td>
		<td align=center height=21 class=tablebody1><%=menpai_setting(3)%></td>
		<td align=center height=21 class=tablebody1><%=menpai_setting(4)%></td>
		<td align=center height=21 class=tablebody1><%=menpai_setting(5)%></td>
		<td align=center height=21 class=tablebody1><%=GetUserTitle(menpai_setting(6))%>(<%=menpai_setting(6)%>级)</td>
		<td align=center height=21 class=tablebody1><%=wumenwupai%></td>	
	</tr>
	<tr>
		<td align=center height=21 class=tablebody2>您的条件</td>
		<td align=center height=21 class=tablebody1><%=mymoney%></td>
		<td align=center height=21 class=tablebody1><%=myuserep%></td>
		<td align=center height=21 class=tablebody1><%=myusercp%></td>
		<td align=center height=21 class=tablebody1><%=mypower%></td>
		<td align=center height=21 class=tablebody1><%=myArticle%></td>
		<td align=center height=21 class=tablebody1><%=myClass%>(<%=GetUserClass(myClass)%>级)</td>
		<td align=center height=21 class=tablebody1><%=mymenpai%></td>		
	</tr>
	<tr>
		<td align=center height=21 class=tablebody2>是否满足</td>
		<td align=center height=21 class=tablebody1><%if int(mymoney)>=int(menpai_setting(1))   then%><font color=blue>√</font><%else%><%CanJoinin=false%><font color=red>×</font><%end if%></td>
		<td align=center height=21 class=tablebody1><%if int(myuserep)>=int(menpai_setting(2))  then%><font color=blue>√</font><%else%><%CanJoinin=false%><font color=red>×</font><%end if%></td>
		<td align=center height=21 class=tablebody1><%if int(myusercp)>=int(menpai_setting(3))  then%><font color=blue>√</font><%else%><%CanJoinin=false%><font color=red>×</font><%end if%></td>
		<td align=center height=21 class=tablebody1><%if int(mypower)>=int(menpai_setting(4))   then%><font color=blue>√</font><%else%><%CanJoinin=false%><font color=red>×</font><%end if%></td>
		<td align=center height=21 class=tablebody1><%if int(myArticle)>=int(menpai_setting(5)) then%><font color=blue>√</font><%else%><%CanJoinin=false%><font color=red>×</font><%end if%></td>
		<td align=center height=21 class=tablebody1><%if int(GetUserClass(myClass))>=int(menpai_setting(6)) then%><font color=blue>√</font><%else%><%CanJoinin=false%><font color=red>×</font><%end if%></td>
		<td align=center height=21 class=tablebody1><%if mymenpai=wumenwupai then%><font color=blue>√</font><%else%><%CanJoinin=false%><font color=red>×</font><%end if%></td>
	</tr>
	<tr><td height=8 class=tablebody2 colspan=8></td></tr>
  	<tr> 
    	<td valign=middle align=center class=tablebody2 height="22">加入本派后</td> 
    	<td valign=middle align=center class=tablebody1 colspan=3><font color=navy> 
<%
	if cint(Menpai_setting(0))=0 then 			'加入 明门正派 
    	response.write "您的金钱-0，经验+0，魅力+0，威望+0"                       
	elseif cint(Menpai_setting(0))=1 then		'加入 社团 
    	response.write "您的金钱-"&int(Menpai_setting(1)*Meipai_inout_rate(0))&"，经验+"&int(menpai_setting(2)*Meipai_inout_rate(1))&"，魅力+"&int(Menpai_setting(3)*Meipai_inout_rate(2))&"，威望+"&int(Menpai_setting(4)*Meipai_inout_rate(3))&""
  	elseif cint(Menpai_setting(0))=2 then		'加入 魔教 
    	response.write "您的金钱+"&int(Menpai_setting(1)*Meipai_inout_rate(4))&"，经验-"&int(menpai_setting(2)*Meipai_inout_rate(5))&"，魅力-"&int(Menpai_setting(3)*Meipai_inout_rate(6))&"，威望-"&int(Menpai_setting(4)*Meipai_inout_rate(7))&""		
	else
		response.write "您的金钱-0，经验+0，魅力+0，威望+0"
	end if
%>                           
  		</font></td>		
  		<td valign=middle align=center class=tablebody2 height="22" colspan=2>您的状态</td> 
  		<td valign=middle align=center class=tablebody1 height="22" colspan=2>
<%
	if mymenpai=wumenwupai then
		response.write "<font color=red>您还没有加入任何门派哟</font> "
	elseif mymenpai<>Groupname then
		response.write mymenpai &"成员"
		CanJoinin=false
	elseif CanManage then
		response.write "本派掌门"
		CanJoinin=false
	else
		response.write "本派弟子"
		CanJoinin=false
		CanExitout=true
	end if
%>
  		</td>                          
  	</tr> 
  	<tr> 
    	<td valign=middle align=center class=tablebody2 height="22">退出本派后</td> 
    	<td valign=middle align=center class=tablebody1 colspan=3><font color=navy> 
<%
	if cint(Menpai_setting(0))=0 then 			'退出 明门正派
    	response.write "您的金钱-0，经验+0，魅力+0，威望+0"                       
	elseif cint(Menpai_setting(0))=1 then		'退出 社团
    	response.write "您的金钱-"&int(Menpai_setting(1)*Meipai_inout_rate(8))&"，经验-"&int(menpai_setting(2)*Meipai_inout_rate(9))&"，魅力-"&int(Menpai_setting(3)*Meipai_inout_rate(10))&"，威望-"&int(Menpai_setting(4)*Meipai_inout_rate(11))&""
  	elseif cint(Menpai_setting(0))=2 then		'退出 魔教
    	response.write "您的金钱-"&int(Menpai_setting(1)*Meipai_inout_rate(12))&"，经验+"&int(menpai_setting(2)*Meipai_inout_rate(13))&"，魅力+"&int(Menpai_setting(3)*Meipai_inout_rate(14))&"，威望+"&int(Menpai_setting(4)*Meipai_inout_rate(15))&""		
	else
		response.write "您的金钱-0，经验+0，魅力+0，威望+0"
	end if
%>                           
  		</font></td>		
  		<td valign=middle align=center class=tablebody2 height="22" colspan=2>您的操作</td> 
  		<td valign=middle align=center class=tablebody1 height="22" colspan=2>
<%
	if CanExitout then
		yourmsg="<font color=red>热烈欢迎本派弟子 "&membername&" 回来</font>"                            
		response.Write "<a href=""z_menpaimember.asp?action=弟子列表&groupname="&htmlencode(groupname)&"""><font color=blue>弟子列表</font></a> | <a href=""z_menpaiadd.asp?menu=3&groupname="&htmlencode(groupname)&""" onclick=""{if(confirm('您确定退出 "&groupname&" 吗?')){return true;}return false;}""><font color=blue>退出本派</font></a> | <a href=z_menpai.asp><font color=blue>门派首页</font></a>"  
	end if
		                                                                                                                                                                                                  
	if CanManage then                                                                                                                                                                                                                                  
		yourmsg="<font color=red>欢迎掌门大驾光临</font>"
		response.Write "<a href=""z_menpaiadd.asp?menu=4&groupname="&htmlencode(groupname)&"""><font color=blue>管理本派</font></a> | <a href=z_menpai.asp><font color=blue>门派首页</font></a>"
	end if                                                                                                                                                                                                                                    
                                                                                                                                                                                                                                  
	if (not Canjoinin) and (not CanExitout) and (not CanManage) then
		if isZangmen then
			yourmsg="<font color=red>热烈欢迎 "&mymenpai&" 掌门人 "&membername&" 光临本派</font>" 
		elseif mymenpai<>wumenwupai then	
			yourmsg="<font color=red>热烈欢迎 "&mymenpai&" 弟子 "&membername&" 光临本派</font>"
		else
			yourmsg="<font color=red>您至少有一项条件不满足，暂时不能加入本派哟</font>"  
		end if	
		response.Write "<a href=""z_menpaiadd.asp?menu=2&groupname="&htmlencode(groupname)&""" onclick=""{if(confirm('您确定加入 "&groupname&" 吗?')){return true;}return false;}""><font color=blue>加入本派</font></a> | <a href=z_menpai.asp><font color=blue>门派首页</font></a>"
	elseif Canjoinin and (not CanManage) then
		if Groupname=wumenwupai then
			response.Write "<a href=z_menpai.asp><font color=blue>门派首页</font></a>"
			yourmsg="<font color=red>江湖险恶，处处要小心哦</font>"
		else
			yourmsg="<font color=red>欢迎江湖侠客 "&membername&" 光临 "&Groupname&"，欢迎您加入本派</font>"
			response.Write "<a href=""z_menpaiadd.asp?menu=2&groupname="&htmlencode(groupname)&""" onclick=""{if(confirm('您确定加入 "&groupname&" 吗?')){return true;}return false;}""><font color=blue>加入本派</font></a> | <a href=z_menpai.asp><font color=blue>门派首页</font></a>"
		end if
	end if                                                                                                                                                                                                                                 
%>                                                                                                                                                                                                                                   
  		</td>                          
  	</tr> 	                               
  	<tr>                                
    	<td valign=middle align=center height=22 class=tablebody1 colspan=8><%=yourmsg%></td>                                                                                                                                                                                                                                   
  	</tr>                                                                                                                                                                                                                                   
</table>                                                                                                                                                                                                                                   
<%                                                                                                                                                                                                                                   
end sub                                                                                                                                                                                                                                  
'-------------------------加入门派----------------------------
sub inmenpai()                                                                                                                                                                                                                                   

	if Groupname="" then 
		Errmsg=Errmsg+"<br>"+"<li>错误的门派参数，请指定相关门派"
		founderr=true
		exit sub		
	end if 	
	if mymenpai<>wumenwupai then
		Errmsg=Errmsg+"<br>"+"<li>您已经是另一个门派的弟子了，要加入本派请先退出原来的门派"
		founderr=true
		exit sub
	elseif int(mymoney)<int(menpai_setting(1)) or int(myuserep)<int(menpai_setting(2)) or int(myusercp)<int(menpai_setting(3)) or int(mypower)<int(menpai_setting(4)) or int(myArticle)<int(menpai_setting(5)) or int(GetUserClass(myClass))<int(menpai_setting(6)) then
		Errmsg=Errmsg+"<br>"+"<li>您至少有一个条件不满足加入本派的要求"
		founderr=true
		exit sub
	end if 		
	                                                                                                                                                                                                                              
    Set rs= Server.CreateObject("ADODB.Recordset")                                                                                                                                                                                                                           
	sql="select usergroup,userwealth,userep,usercp,userpower from [user] where username='"&membername&"'"                                                                                                                                                                                                                                   
	rs.open sql,conn,1,3                                                                                                                                                                                                                                  
	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br>"+"<li>您怎么会跑到这里来了"
		founderr=true
		exit sub
	else
		sucmsg=sucmsg+"<br>"+"<li>您成功加入了 "&groupname		
		if cint(menpai_setting(0))=0 then       	'名门正派 
			rs(0)=Groupname
			rs.update
			sucmsg=sucmsg+"<br>"+"<li>您的金钱-0，经验+0，魅力+0，威望+0"
		elseif cint(menpai_setting(0))=1 then       '社团
			rs(0)=Groupname
			rs(1)=rs(1)-int(menpai_setting(1)*Meipai_inout_rate(0))
			rs(2)=rs(2)+int(menpai_setting(2)*Meipai_inout_rate(1))
			rs(3)=rs(3)+int(menpai_setting(3)*Meipai_inout_rate(2))
			rs(4)=rs(4)+int(menpai_setting(4)*Meipai_inout_rate(3))
			rs.update
			sucmsg=sucmsg+"<br>"+"<li>您的金钱-"&int(Menpai_setting(1)*Meipai_inout_rate(0))&"，经验+"&int(menpai_setting(2)*Meipai_inout_rate(1))&"，魅力+"&int(Menpai_setting(3)*Meipai_inout_rate(2))&"，威望+"&int(Menpai_setting(4)*Meipai_inout_rate(3))
		elseif cint(menpai_setting(0))=2 then       '魔教           
			rs(0)=Groupname
			rs(1)=rs(1)+int(menpai_setting(1)*Meipai_inout_rate(4))        
			rs(2)=rs(2)-int(menpai_setting(2)*Meipai_inout_rate(5))   
			rs(3)=rs(3)-int(menpai_setting(3)*Meipai_inout_rate(6))
			rs(4)=rs(4)-int(menpai_setting(4)*Meipai_inout_rate(7))
			rs.update
			sucmsg=sucmsg+"<br>"+"<li>您的金钱+"&int(Menpai_setting(1)*Meipai_inout_rate(4))&"，经验-"&int(menpai_setting(2)*Meipai_inout_rate(5))&"，魅力-"&int(Menpai_setting(3)*Meipai_inout_rate(6))&"，威望-"&int(Menpai_setting(4)*Meipai_inout_rate(7))&""
		else
			Errmsg=Errmsg+"<br>"+"<li>错误的门派参数，请指定相关的门派属性参数"
			founderr=true
			exit sub		
		end if
	end if
	
	call menpai_suc()
end sub                                                                                                                                                                                                                                    	  
'---------------------------退出门派------------------------------
sub exitmenpai()                                                                                                                                                                                                                                   

	if Groupname="" then 
		Errmsg=Errmsg+"<br>"+"<li>错误的门派参数，请指定相关门派"
		founderr=true
		exit sub
	end if
	
	if cint(Menpai_setting(0))=0 then 
	
	elseif cint(Menpai_setting(0))=1 then 
		if int(menpai_setting(2)*Meipai_inout_rate(9))>myuserep or int(Menpai_setting(3)*Meipai_inout_rate(10))>myusercp or int(Menpai_setting(4)*Meipai_inout_rate(11))>mypower then                                                                                                                                                                                                      
			Errmsg=Errmsg+"<br>"+"<li>您还没有足够的资本脱离本派"
			founderr=true
		end if
	elseif cint(Menpai_setting(0))=2 then
		if int(Menpai_setting(1)*Meipai_inout_rate(12))>mymoney then
			Errmsg=Errmsg+"<br>"+"<li>您还没有足够的资本脱离本派"
		end if
	else
		Errmsg=Errmsg+"<br>"+"<li>错误的门派参数，请指定相关门派属性"
		founderr=true
	end if
		
	if founderr then exit sub
		
	Groupname=wumenwupai  
                                                                                                                                                                                                                   
	Set rs= Server.CreateObject("ADODB.Recordset")                                                                                                                                                                                                                                   
	sql="select usergroup,userwealth,userep,usercp,userpower from [user] where username='"&membername&"'"                                                                                                                                                                                                                                   
	rs.open sql,conn,1,3                                                                                                                                                                                                                                  
                                                                                                                                                                                                                   
	if cint(Menpai_setting(0))=0 then   		'名门正派                                                                                                                                                                                                                        
		rs("usergroup")=groupname                                                                                                                                                                                                                           
		rs.update     
		sucmsg=sucmsg+"<br>"+"<li>您的金钱+0，经验-0，魅力-0，威望-0"                                                                                                                                                                                                                      
	elseif cint(Menpai_setting(0))=1 then     	'社团	                                                                                                                                                                                                                              
		rs(0)=Groupname
		rs(1)=rs(1)-int(menpai_setting(1)*Meipai_inout_rate(8))
		rs(2)=rs(2)-int(menpai_setting(2)*Meipai_inout_rate(9))
		rs(3)=rs(3)-int(menpai_setting(3)*Meipai_inout_rate(10))
		rs(4)=rs(4)-int(menpai_setting(4)*Meipai_inout_rate(11))                                                                                                                                                                                                                         
		rs.update 
		sucmsg=sucmsg+"<br>"+"<li>您的金钱+"&int(Menpai_setting(1)*Meipai_inout_rate(8))&"，经验-"&int(menpai_setting(2)*Meipai_inout_rate(9))&"，魅力-"&int(Menpai_setting(3)*Meipai_inout_rate(10))&"，威望-"&int(Menpai_setting(4)*Meipai_inout_rate(11))&""                                                                                                                                                                                                                         
	elseif cint(Menpai_setting(0))=2 then   	'魔教
		rs(0)=Groupname
		rs(1)=rs(1)-int(menpai_setting(1)*Meipai_inout_rate(12))        
		rs(2)=rs(2)+int(menpai_setting(2)*Meipai_inout_rate(13))   
		rs(3)=rs(3)+int(menpai_setting(3)*Meipai_inout_rate(14))
		rs(4)=rs(4)+int(menpai_setting(4)*Meipai_inout_rate(15))                                                                                                                                                                                                                         
		rs.update  
		sucmsg=sucmsg+"<br>"+"<li>您的金钱-"&int(Menpai_setting(1)*Meipai_inout_rate(12))&"，经验+"&int(menpai_setting(2)*Meipai_inout_rate(13))&"，魅力+"&int(Menpai_setting(3)*Meipai_inout_rate(14))&"，威望+"&int(Menpai_setting(4)*Meipai_inout_rate(15))&""                                                                                                                                                                                                                        
	end if  
	rs.close
	set rs=nothing                                                                                                                                                                                                                        
	call menpai_suc()                                                                                                                                                                                                                                
end sub                                                                                                                                                                                                                                   
'---------------------------门派管理====管理页面------------------------------
sub manage()                                                                                                                                                                                           
	if not CanManage then
		Errmsg=Errmsg+"<br>"+"<li>您不是本派掌门，没有管理本派的权限"
		founderr=true
		exit sub	
	end if
	if Groupname="" then 
		Errmsg=Errmsg+"<br>"+"<li>错误的门派参数，请指定相关门派参数"
		founderr=true
		exit sub
	end if	 

%>   
<table cellspacing=1 align=center class=tableborder1 width=<%=Forum_body(12)%>>                                                                                                                                                                              
  	<tr>                                                                                                                                                                                           
    	<td valign=middle align=center height=24 class=tablebody2 colspan="4">门派管理</td>                                                                                                                                                                                          
  	</tr> 
	<form name="form" method="post" action="z_menpaiadd.asp?menu=6">                                                                                                                                                                                            
  	<tr>                                                                                                                                                                                           
    	<th valign=middle align=left height=24 colspan="4">&nbsp;★&nbsp; 门派基本设置</th>                                                                                                                                                                                          
  	</tr> 
	<tr><td height=8 class=tablebody2 colspan="4"></td></tr>                                                                                                                                                                                        
    <tr>                                                                                                                                                                                          
		<td valign=middle align=left class=tablebody2 height="22"><b>掌门须知:</b></td>                                                                                                                                                                                         
		<td valign=middle align=left class=tablebody1 height="22" colspan="3">
		 	<li><font color="#FF9966">欢迎 <font color=blue><%=Groupname%></font> 掌门人 <font color=red><%=membername%></font> 进入门派管理</font>
			<li>门派的最终解释权属于管理员
			<li>如果掌门人不负责任，管理员有权撤消其掌门人的职务 
			<li>下面是您所管理的门派的基本情况
		</td> 
  	</tr> 	   
	<tr><td height=8 class=tablebody2 colspan="4"></td></tr> 
  	<tr> 
		<td valign=middle align=left class=tablebody2 height="22"><b>门派名称:</b><br> 门派名称申请之后掌门不能修改</td> 
		<td valign=middle align=left class=tablebody1 height="22" colspan="3">&nbsp;<INPUT name=newGroupName size="20" maxlength="20" disabled value=<%=GroupName%>> <font color=red>*</font> (不要超过12个英文字母或6个汉字)<INPUT name=GroupName size="20" maxlength="20" type=hidden value=<%=GroupName%>></td>
  	</tr>
  	<tr> 
		<td valign=middle align=left class=tablebody2 height="22"><b>门派性质:</b><br> 门派性质申请之后掌门不能修改</td>
		<td valign=middle align=left class=tablebody1 height="22" colspan="3">
			<input type=radio name=newMenpaiAttri value=0 disabled <%if cint(menpai_setting(0))=0 then%>checked<%end if%>>明门正派&nbsp;
			<input type=radio name=newMenpaiAttri value=1 disabled <%if cint(menpai_setting(0))=1 then%>checked<%end if%>>社团&nbsp;
			<input type=radio name=newMenpaiAttri value=2 disabled <%if cint(menpai_setting(0))=2 then%>checked<%end if%>>魔教
			<input type=hidden name=MenpaiAttri value=<%=cint(menpai_setting(0))%>>
		</td>
  	</tr>
	<tr><td height=8 class=tablebody2 colspan="4"></td></tr>
    <tr>
		<td valign=middle align=left class=tablebody2 height="22"><b>说明:</b></td>                                                                                                                                                                                         
		<td valign=middle align=left class=tablebody1 height="22" colspan="3">
		 	<li>金钱,魅力,经验,威望,文章数,用户等级 都不能超过您自已的当前值
			<li>金钱,魅力,经验,威望,文章数,用户等级 是加入本派的必须满足的最小值
			<li>金钱,魅力,经验,威望,文章数,用户等级 也是加入和退出门派增加/扣除相关值时的基数
			<li>明门正派不受以上限制
			<li>掌门可以根据实际情况修改下面各项，修改好之后 提交 即可			
		</td> 
  	</tr> 
    <tr>                                                                                                                                                                                         
		<td valign=middle align=left class=tablebody2 height="22" width="25%"><b>金钱:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=userwealth value=<%=menpai_setting(1)%> <%if cint(menpai_setting(0))=0 then%> disabled<%end if%>></td>                                                                                                                                                                                        
		<td valign=middle align=right class=tablebody2 height="22">您的金钱:&nbsp;</td>
		<td valign=middle align=left class=tablebody1 height="22" width="20%">&nbsp;<%=mymoney%></td>
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>经验:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=userep value=<%=menpai_setting(2)%> <%if cint(menpai_setting(0))=0 then%> disabled<%end if%>></td>
		<td valign=middle align=right class=tablebody2 height="22">您的经验:&nbsp;</td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<%=myusercp%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>魅力:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=usercp value=<%=menpai_setting(3)%> <%if cint(menpai_setting(0))=0 then%> disabled<%end if%>></td> 
		<td valign=middle align=right class=tablebody2 height="22">您的魅力:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<%=myuserep%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>威望:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=userpower value=<%=menpai_setting(4)%> <%if cint(menpai_setting(0))=0 then%> disabled<%end if%>></td> 
		<td valign=middle align=right class=tablebody2 height="22">您的威望:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<%=mypower%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>文章数:</b><br>加入本派必须达到的发帖数</td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=userarticle value=<%=menpai_setting(5)%> <%if cint(menpai_setting(0))=0 then%> disabled<%end if%>></td> 
		<td valign=middle align=right class=tablebody2 height="22">您的帖数:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<%=myarticle%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>用户等级:</b><br>加入本派必须具备的等级(1～<%=maxtitleid%>)</td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=userclass value=<%=menpai_setting(6)%> <%if cint(menpai_setting(0))=0 then%> disabled<%end if%>>&nbsp;<%=GetUserTitle(Menpai_setting(6))%>(<%=Menpai_setting(6)%>级)</td> 
		<td valign=middle align=right class=tablebody2 height="22">您的等级:&nbsp;</td>                                                                                                                                                                                     
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<%=myclass%>（<%=GetUserClass(myclass)%>级）</td>                                                                                                                                                                                      
  	</tr>
	<tr>
		<td valign=middle align=center class=tablebody1 height=22 colspan="4"><font color="#CC9933">
		如果门派性质是明门正派，则以上6项(金钱,魅力,经验,威望,文章数,用户等级)不必填写</font>
		</td>
	</tr> 
	<tr><td height=8 class=tablebody2 colspan="4"></td></tr>                                                                                                                                                                                       
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>本派简介:</b><br>对您的门派作简单的介绍，在您的门派资料中会显示出来</td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 colspan="3" height="22">&nbsp;<textarea name=menpaireadme rows="4" cols="60"><%=menpai_setting(7)%></textarea>		(不要超过100个英文或50个汉字)</td>                                                                                                                                                                                      
  	</tr>  
	<tr><td height=8 class=tablebody2 colspan="4"></td></tr>                                                                                                                                                                                       
  	<tr>                                                                                                                                                                                       
    	<td valign=middle align=center class=tablebody1 colspan="4" height="22"><INPUT type=submit value=更新 name=submit></td>                                                                                                                                                                                       
  	</tr>                                                                                                                                                                              
  	</form>                                                                                                                                                                                    
    <tr><td height=15 class=tablebody2 colspan="4"></td></tr>                                                                                                                                                                                                   
	<form name="form1" method="post" action="z_menpaimember.asp">                                                                                                                                                                               
	<INPUT name=Groupname size="20" value="<%=groupname%>" type=hidden>                                                                                                                                                                          
	<tr>                                                                                                                                                                                         
		<th valign=middle align=left height=24 colspan="4" >&nbsp;◆&nbsp;<b>门派人员快捷管理</b></th>                                                                                                                                                                                         
	</tr>                                                                                                                                                                               
	<tr>                                                                                                                                                                                         
		<td valign=middle align=left class=tablebody2 height="22"><b>本派弟子</b><br>请输入本派弟子的用户名</td>                                                                                                                                                                                         
		<td valign=middle align=left class=tablebody1 colspan="3">
		&nbsp;<INPUT name=Susername type=text>&nbsp;&nbsp;
		<INPUT type=submit value=查找 name=action>&nbsp;&nbsp;
		<INPUT type=checkbox value=1 name=search checked>完全匹配用户名&nbsp;&nbsp;
		<INPUT type=submit value=弟子列表 name=action>&nbsp;&nbsp;
	</tr>
	</form>  
	<form name="form3" method="post" action="z_menpaiadd.asp?menu=5">                                                                                                                                                                               
	<INPUT name=Groupname size="20" value="<%=groupname%>" type=hidden> 	
  	<tr>                                                                                                                                                                                       
    	<td valign=middle align=left class=tablebody2 height="22"><b>清理门户</b><br>请输入您想踢出本派的用户名</td>
    	<td valign=middle align=left class=tablebody1 colspan="3">&nbsp;<INPUT name=Tusername type=text>&nbsp;&nbsp;                                                                                                                                                                        
      		<INPUT type=submit value=清理 name=submit onclick="{if(confirm('您确定执行踢人操作吗?')){this.document.form3.submit();return true;}return false;}">
		</td>                                                                                                                                                                                         
  	</tr>                                                                                                                                                                       
  	</form>  
  	<tr><td height=8 class=tablebody2 colspan="4"></td></tr>                                                                                                                                                                    
	<tr>                                                                                                                                                                                           
		<th valign=middle align=left height=24 colspan="4">&nbsp;■&nbsp;注销本派</th>                                                                                                                                                                                          
	</tr>                                                                                                                                                                                 
  	<form name="form2" method="post" action="z_menpaiadd.asp?menu=9">                                                                                                                     
  	<INPUT name=Groupname size="20" value="<%=Groupname%>" type=hidden>                                                                                                                                                                     
	<tr>  
		<td valign=middle align=left class=tablebody2 height="22" colspan="2"><b>注销本派</b><br>该操作将会删除该门派，请考虑清楚后进行操作</td>
		<td valign=middle align=center class=tablebody1 colspan="2">
		  	<INPUT type=submit value=注销 name=submit onclick="{if(confirm('您确定要注销本派吗?')){this.document.form2.submit();return true;}return false;}">                                                                                                                                                                  
		</td>                                                                                                                                                                                          
	</tr>                                                                                                                                                                   
  </form>                                                                                                                                                                                              
 </table>                                                                                                                                                                                            
<%                                                                                                                                                                                                                             
end sub                                                                                                                                                                                                                             
'----------------------------------清理门户-----------------------------------
sub clearmember()   
	if not CanManage then
		Errmsg=Errmsg+"<br>"+"<li>您不是本派掌门，没有管理本派的权限"
		founderr=true
		exit sub	
	end if
	                                                                                                                                                                                            
	dim Tusername                                                                                                                                                                                                                                   
	Groupname=checkstr(trim(Request.form("Groupname")))
	Tusername=checkstr(trim(Request.form("Tusername")))
	
	if Groupname="" or isnull(Groupname) then
		Errmsg=Errmsg+"<br>"+"<li>错误的门派参数，请指定相关门派参数"
		founderr=true	
	end if
	if Tusername="" or isnull(Tusername) then
		Errmsg=Errmsg+"<br>"+"<li>请输入您想踢出本派的用户名"
		founderr=true	
	end if
	if founderr then exit sub
		
	Set rs= Server.CreateObject("ADODB.Recordset")                                                                                                                                                                                                                                   
	sql="select * from [user] where username='"&Tusername&"'"                                                                                                                                                                                                                                   
	rs.open sql,conn,1,3                                                                                              
	if rs.eof and rs.bof then                                                                                             
		Errmsg=Errmsg+"<br>"+"<li>您输入的用户不存在"
		founderr=true
		exit sub
	elseif rs("usergroup")<>Groupname then                                                                                             
		Errmsg=Errmsg+"<br>"+"<li>您输入的用户不是本派弟子"
		founderr=true
		exit sub                                                                                            
	else                                                                                             
		rs("usergroup")=wumenwupai
		rs.update                                                                                             
	end if                                                                                                                                                                                            
	rs.close                                                                                             
	set rs=nothing  
	
	call sendmessage(Tusername,Groupname)	 '发送清理短信息给用户
	Sucmsg=Sucmsg+"<br>"+"<li>您已经成功把[<font color=red>"&Tusername&"</font>]清理出本派"
	Sucmsg=Sucmsg+"<br>"+"<li>已经发送了清理通知给该用户" 	 
	call menpai_suc()	                                                                                                                                                                                                                          
end sub                                                                                                                                                                                                                             
'-------------------------------保存门派资料的修改---------------------------
sub changegroup()  
	if not CanManage then
		Errmsg=Errmsg+"<br>"+"<li>您不是本派掌门，没有管理本派的权限"
		founderr=true
		exit sub	
	end if
	                                                                                                                                                   
	dim MenpaiAttri,userwealth,usercp,userep,userpower,userclass,userartitle
	'GroupName=checkstr(trim(Request.form("GroupName")))
	'newGroupName=checkstr(trim(Request.form("newGroupName")))
	GroupName=checkstr(trim(Request.form("GroupName")))
	MenpaiAttri=Request.form("MenpaiAttri")		'MenpaiAttri=Request.form("newMenpaiAttri")
	
	if GroupName="" then		'保留这个功能
		Errmsg=Errmsg+"<br>"+"<li>请输入门派名称"
		founderr=true
	elseif strLength(GroupName)>12 then
		Errmsg=Errmsg+"<br>"+"<li>门派名称不能超过12个英文字母(6个汉字)"
		founderr=true		
	end if		
	if cint(MenpaiAttri)<>0 then
		if not isInteger(Request.form("userwealth")) then
			Errmsg=Errmsg+"<br>"+"<li>非明门正派必须填写金钱值或者您填入的不是有效的数值"
			founderr=true
		elseif int(Request.form("userwealth"))>mymoney then
			Errmsg=Errmsg+"<br>"+"<li>您输入的金钱有效值大于您自己的当前值！Request.form(""userwealth"")="&Request.form("userwealth")&"  mymoney="&mymoney&"==="&(Request.form("userwealth")>mymoney)
			founderr=true			
		else
			userwealth=int(Request.form("userwealth"))
		end if	
		if not isInteger(Request.form("userep")) then
			Errmsg=Errmsg+"<br>"+"<li>非明门正派必须填写经验值或者您填入的不是有效的数值"
			founderr=true
		elseif int(Request.form("userep"))>myuserep then
			Errmsg=Errmsg+"<br>"+"<li>您输入的经验有效值大于您自己的当前值！"
			founderr=true			
		else 
			userep=int(Request.form("userep"))	
		end if
		if not isInteger(Request.form("usercp")) then
			Errmsg=Errmsg+"<br>"+"<li>非明门正派必须填写魅力值或者您填入的不是有效的数值"
			founderr=true
		elseif int(Request.form("usercp"))>myusercp then
			Errmsg=Errmsg+"<br>"+"<li>您输入的魅力有效值大于您自己的当前值！"
			founderr=true			
		else 
			usercp=int(Request.form("usercp"))	
		end if
		if not isInteger(Request.form("userpower")) then
			Errmsg=Errmsg+"<br>"+"<li>非明门正派必须填写威望值或者您填入的不是有效的数值"
			founderr=true
		elseif int(Request.form("userpower"))>mypower then
			Errmsg=Errmsg+"<br>"+"<li>您输入的威望有效值大于您自己的当前值！"&mypower
			founderr=true			
		else 
			userpower=int(Request.form("userpower"))	
		end if
		if not isInteger(Request.form("userarticle")) then
			Errmsg=Errmsg+"<br>"+"<li>非明门正派必须填写文章数或者您填入的不是有效的数值"
			founderr=true
		elseif int(Request.form("userarticle"))>myarticle then
			Errmsg=Errmsg+"<br>"+"<li>您输入的文章有效值大于您自己的当前值！"
			founderr=true			
		else 
			userartitle=int(Request.form("userarticle"))	
		end if
		if not isInteger(Request.form("userclass")) then
			Errmsg=Errmsg+"<br>"+"<li>非明门正派必须填写威望值或者您填入的不是有效的数值"
			founderr=true
		elseif int(Request.form("userclass"))>GetUserClass(myclass) then
			Errmsg=Errmsg+"<br>"+"<li>您输入的等级有效值大于您自己的当前值！"
			founderr=true
		elseif int(Request.form("userclass"))<1 or int(Request.form("userclass"))>int(maxtitleid) then
			Errmsg=Errmsg+"<br>"+"<li>等级值不能小于1或大于"&maxtitleid
			founderr=true						
		else 
			userclass=int(Request.form("userclass"))	
		end if
	else
		userwealth=0
		userep=0
		usercp=0
		userpower=0
		userartitle=0
		userclass=1
	end if
	
	dim menpaireadme
	menpaireadme=Request.form("menpaireadme")	
	if strLength(menpaireadme)>100 then
		Errmsg=Errmsg+"<br>"+"<li>门派简介不能超过100个英文或50个汉字"
		founderr=true		
	end if
			
	if founderr then exit sub
																																																						   
	set rs= server.createobject ("adodb.recordset")                                                                                                                                                                                                                                   
	sql="select * from [GroupName] where GroupName='"&GroupName&"'"                                                                                                                                                                                                                                   
	rs.Open sql,conn,1,3  
	 
	if rs.bof and rs.eof then   	'该功能以后扩展可以修改门派名字是用                                                                            
		Errmsg=Errmsg+"<br>"+"<li>该门派不存在啊"
		founderr=true
		rs.close
		set rs=nothing
		exit sub
	else                                  
		'rs("Groupname")=GroupName
		rs("Menpai_setting")=MenpaiAttri & "|" & userwealth & "|" & userep & "|" & usercp & "|" & userpower & "|" & userartitle & "|" & userclass & "|" & menpaireadme
		rs.update  
		'conn.execute("update [user] set usergroup='"&GroupName&"',userwealth=userwealth-"&new_menpai_setting(6)&" where username='"&membername&"'")	                                            
	end if
	Sucmsg=Sucmsg+"<br>"+"<li>门派资料修改成功" 	 
	call menpai_suc()
end sub                                                                                                                                                                                                                             
'----------------------------注销门派-------------------------------
sub logout()
                                                                                              
	if not CanManage then
		Errmsg=Errmsg+"<br>"+"<li>您不是本派掌门，没有管理本派的权限"
		founderr=true
		exit sub	
	end if
	if Groupname="" then 
		Errmsg=Errmsg+"<br>"+"<li>错误的门派参数，请指定相关门派参数"
		founderr=true
		exit sub
	elseif Groupname=wumenwupai then
		Errmsg=Errmsg+"<br>"+"<li>本派是收留所有无门无派的人的特殊门派，不可注销"
		founderr=true
		exit sub	
	end if
	                                                                                  
    conn.execute("update [GroupName] set state=2,Zangmen='-弃帮跑路-' where GroupName='"&GroupName&"'")
	conn.execute("update [user] set usergroup='"&wumenwupai&"' where usergroup='"&GroupName&"'")

	Sucmsg=Sucmsg+"<br>"+"<li>您已经注销了本派" 	 
	call menpai_suc()	
end sub                                                                                                                                                                                                                             
'-----------------------------申请新门派==界面------------------------------
sub newmenpai()       
	if isZangmen then
		Errmsg=Errmsg+"<br>"+"<li>您已经是 "&Groupname&" 掌门，怎么还来创派呢"
		founderr=true
		exit sub
	end if
	if mymenpai<>wumenwupai then
		Errmsg=Errmsg+"<br>"+"<li>您已经是 "&Groupname&" 弟子，不能创建新门派"
		founderr=true
		exit sub		
	end if
	
	dim new_menpai_setting
	new_menpai_setting=split(conn.execute("select new_menpai_setting from [menpai]")(0),",")
		
	if int(mymoney)<(int(new_menpai_setting(0))+int(new_menpai_setting(6))) or int(myusercp)<int(new_menpai_setting(1)) or int(myuserep)<int(new_menpai_setting(2)) or int(mypower)<int(new_menpai_setting(3)) or int(myArticle)<int(new_menpai_setting(4)) or int(GetUserClass(myClass))<int(new_menpai_setting(5)) then
		Errmsg=Errmsg+"<br>"+"<li>您至少有一个条件不满足创派要求"
		founderr=true
		exit sub
	end if	
%>                                                                                             
<table cellspacing=1 align=center class=tableborder1 width=<%=Forum_body(12)%>>                                                                                                                                                                              
	<form name="form" method="post" action="z_menpaiadd.asp?menu=7">                                                                                                                                                                                            
  	<tr>                                                                                                                                                                                           
    	<th valign=middle align=center height=24 colspan="4">新门派申请</th>                                                                                                                                                                                          
  	</tr> 
	<tr><td height=8 class=tablebody2 colspan="4"></td></tr>                                                                                                                                                                                        
    <tr>                                                                                                                                                                                          
		<td valign=middle align=left class=tablebody2 height="22"><b>创派须知:</b></td>                                                                                                                                                                                         
		<td valign=middle align=left class=tablebody1 height="22" colspan="3">
		 	<li>恭喜您，您具备了创建门派的条件 
			<li>门派创建成功之后，等待管理员的审核，审核通过之后，本门派正式成立 
			<li>门派正式成立之后，您将会是本派的掌门人 
			<li>如果掌门人不负责任，管理员有权撤消其掌门人的职务 
			<li>创派费用<font color=red><%=new_menpai_setting(6)%></font>元  
			<li>请认真填写下面的门派资料
		</td> 
  	</tr> 	   
	<tr><td height=8 class=tablebody2 colspan="4"></td></tr> 
  	<tr> 
		<td valign=middle align=left class=tablebody2 height="22"><b>门派名称:</b><br> 门派名称申请之后掌门不能修改</td> 
		<td valign=middle align=left class=tablebody1 height="22" colspan="3">&nbsp;<INPUT name=groupname size="20" maxlength="20"> <font color=red>*</font> (不要超过12个英文字母或6个汉字)</td>
  	</tr>
  	<tr> 
		<td valign=middle align=left class=tablebody2 height="22"><b>门派性质:</b><br> 门派性质申请之后掌门不能修改</td>
		<td valign=middle align=left class=tablebody1 height="22" colspan="3">
			<input type=radio name=menpaiattri value=0 checked>明门正派&nbsp;
			<input type=radio name=menpaiattri value=1>社团&nbsp;
			<input type=radio name=menpaiattri value=2>魔教<br>
		</td>
  	</tr>
	<tr><td height=8 class=tablebody2 colspan="4"></td></tr>
    <tr>
		<td valign=middle align=left class=tablebody2 height="22"><b>说明:</b></td>                                                                                                                                                                                         
		<td valign=middle align=left class=tablebody1 height="22" colspan="3">
		 	<li>金钱,魅力,经验,威望,文章数,用户等级 都不能超过您自已的当前值
			<li>金钱,魅力,经验,威望,文章数,用户等级 是加入本派的必须满足的最小值
			<li>金钱,魅力,经验,威望,文章数,用户等级 也是加入和退出门派增加/扣除相关值时的基数
			<li>明门正派不受以上限制
			<li>以下属性掌门可以在门派管理中修改			
		</td> 
  	</tr> 
    <tr>                                                                                                                                                                                         
		<td valign=middle align=left class=tablebody2 height="22" width="25%"><b>金钱:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=userwealth value=""></td>                                                                                                                                                                                        
		<td valign=middle align=right class=tablebody2 height="22">您的金钱:&nbsp;</td>
		<td valign=middle align=left class=tablebody1 height="22" width="20%">&nbsp;<%=mymoney%></td>
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>经验:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=userep value=""></td>
		<td valign=middle align=right class=tablebody2 height="22">您的经验:&nbsp;</td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<%=myusercp%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>魅力:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=usercp value=""></td> 
		<td valign=middle align=right class=tablebody2 height="22">您的魅力:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<%=myuserep%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>威望:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=userpower value=""></td> 
		<td valign=middle align=right class=tablebody2 height="22">您的威望:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<%=mypower%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>文章数:</b><br>加入本派必须达到的发帖数</td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=userarticle value=""></td> 
		<td valign=middle align=right class=tablebody2 height="22">您的帖数:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<%=myarticle%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>用户等级:</b><br>加入本派必须具备的等级(1～<%=maxtitleid%>)</td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<INPUT name=userclass value=""></td> 
		<td valign=middle align=right class=tablebody2 height="22">您的等级:&nbsp;</td>                                                                                                                                                                                     
		<td valign=middle align=left class=tablebody1 height="22">&nbsp;<%=myclass%>（<%=GetUserClass(myclass)%>级）</td>                                                                                                                                                                                      
  	</tr>
	<tr>
		<td valign=middle align=center class=tablebody1 height=22 colspan="4"><font color="#CC9933">
		如果门派性质是明门正派，则以上6项(金钱,魅力,经验,威望,文章数,用户等级)不必填写</font>
		</td>
	</tr> 
	<tr><td height=8 class=tablebody2 colspan="4"></td></tr>                                                                                                                                                                                       
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody2 height="22"><b>本派简介:</b><br>对您的门派作简单的介绍，在您的门派资料中会显示出来</td>                                                                                                                                                                                        
		<td valign=middle align=left class=tablebody1 colspan="3" height="22">&nbsp;<textarea name=menpaireadme rows="4" cols="60"></textarea>		(不要超过100个英文或50个汉字)</td>                                                                                                                                                                                      
  	</tr>  
	<tr><td height=8 class=tablebody2 colspan="4"></td></tr>                                                                                                                                                                                       
  	<tr>                                                                                                                                                                                       
    	<td valign=middle align=center class=tablebody1 colspan="4" height="22"><INPUT type=submit value=申请新门派 name=submit></td>                                                                                                                                                                                       
  	</tr>                                                                                                                                                                              
  	</form>                                                                                                                                                                                    
</table>                                                                                                                                                                                           
<%                                             
end sub                                                                                                                                                                                                                            

'------------------------------------新门派申请====提交申请--------------------------------
sub applynewmenpai() 
	dim MenpaiAttri,userwealth,usercp,userep,userpower,userclass,userartitle
	GroupName=checkstr(trim(Request.form("groupname")))
	MenpaiAttri=Request.form("menpaiattri")
	
	if GroupName="" then
		Errmsg=Errmsg+"<br>"+"<li>请输入门派名称"
		founderr=true
	elseif strLength(GroupName)>12 then
		Errmsg=Errmsg+"<br>"+"<li>门派名称不能超过12个英文字母(6个汉字)"
		founderr=true		
	end if		
	if cint(MenpaiAttri)<>0 then
		if not isInteger(Request.form("userwealth")) then
			Errmsg=Errmsg+"<br>"+"<li>非明门正派必须填写金钱值或者您填入的不是有效的数值"
			founderr=true
		elseif int(Request.form("userwealth"))>mymoney then
			Errmsg=Errmsg+"<br>"+"<li>您输入的金钱有效值大于您自己的当前值！Request.form(""userwealth"")="&Request.form("userwealth")&"  mymoney="&mymoney&"==="&(Request.form("userwealth")>mymoney)
			founderr=true			
		else
			userwealth=int(Request.form("userwealth"))
		end if	
		if not isInteger(Request.form("userep")) then
			Errmsg=Errmsg+"<br>"+"<li>非明门正派必须填写经验值或者您填入的不是有效的数值"
			founderr=true
		elseif int(Request.form("userep"))>myuserep then
			Errmsg=Errmsg+"<br>"+"<li>您输入的经验有效值大于您自己的当前值！"
			founderr=true			
		else 
			userep=int(Request.form("userep"))	
		end if
		if not isInteger(Request.form("usercp")) then
			Errmsg=Errmsg+"<br>"+"<li>非明门正派必须填写魅力值或者您填入的不是有效的数值"
			founderr=true
		elseif int(Request.form("usercp"))>myusercp then
			Errmsg=Errmsg+"<br>"+"<li>您输入的魅力有效值大于您自己的当前值！"
			founderr=true			
		else 
			usercp=int(Request.form("usercp"))	
		end if
		if not isInteger(Request.form("userpower")) then
			Errmsg=Errmsg+"<br>"+"<li>非明门正派必须填写威望值或者您填入的不是有效的数值"
			founderr=true
		elseif int(Request.form("userpower"))>mypower then
			Errmsg=Errmsg+"<br>"+"<li>您输入的威望有效值大于您自己的当前值！"
			founderr=true			
		else 
			userpower=int(Request.form("userpower"))	
		end if
		if not isInteger(Request.form("userarticle")) then
			Errmsg=Errmsg+"<br>"+"<li>非明门正派必须填写文章数或者您填入的不是有效的数值"
			founderr=true
		elseif int(Request.form("userarticle"))>myarticle then
			Errmsg=Errmsg+"<br>"+"<li>您输入的文章有效值大于您自己的当前值！"
			founderr=true			
		else 
			userartitle=int(Request.form("userarticle"))	
		end if
		if not isInteger(Request.form("userclass")) then
			Errmsg=Errmsg+"<br>"+"<li>非明门正派必须填写威望值或者您填入的不是有效的数值"
			founderr=true
		elseif int(Request.form("userclass"))>GetUserClass(myclass) then
			Errmsg=Errmsg+"<br>"+"<li>您输入的等级有效值大于您自己的当前值！"
			founderr=true
		elseif int(Request.form("userclass"))<1 or int(Request.form("userclass"))>int(maxtitleid) then
			Errmsg=Errmsg+"<br>"+"<li>等级值不能小于1或大于"&maxtitleid
			founderr=true						
		else 
			userclass=int(Request.form("userclass"))	
		end if
	else
		userwealth=0
		userep=0
		usercp=0
		userpower=0
		userartitle=0
		userclass=1
	end if
	
	dim menpaireadme
	menpaireadme=Request.form("menpaireadme")	
	if strLength(menpaireadme)>100 then
		Errmsg=Errmsg+"<br>"+"<li>门派简介不能超过100个英文或50个汉字"
		founderr=true		
	end if
			
	if founderr then exit sub
	
	dim new_menpai_setting		'new_menpai_setting(6)   创派收费
	new_menpai_setting=split(conn.execute("select new_menpai_setting from [menpai]")(0),",")                         
                           
	set rs= server.createobject ("adodb.recordset")                                                                                                                                                                                                                                   
	sql="select * from [GroupName] where groupname='"&groupname&"'"                                                                                                                                                                                                                                   
	rs.Open sql,conn,1,3                                                                                                                                                                                                            
	if not(rs.bof and rs.eof) then                                                                               
		Errmsg=Errmsg+"<br>"+"<li>您输入的门派名字已经存在"
		founderr=true
		rs.close
		set rs=nothing
		exit sub
	else                                  
		rs.addnew                                                                                                                                                          
		rs("Groupname")=GroupName
		rs("Zangmen")=membername
		rs("Addtime")=date()
		rs("Menpai_setting")=MenpaiAttri & "|" & userwealth & "|" & userep & "|" & usercp & "|" & userpower & "|" & userartitle & "|" & userclass & "|" & menpaireadme
		if cint(new_menpai_setting(7))=1 then
			 rs("State")=3		'等待审核
		else
			 rs("State")=0		'不需要审核，直接开通
		end if 
		rs.update  
		conn.execute("update [user] set usergroup='"&GroupName&"',userwealth=userwealth-"&new_menpai_setting(6)&" where username='"&membername&"'")	                                            
	end if
	if cint(new_menpai_setting(7))=1 then
		Sucmsg=Sucmsg+"<br>"+"<li>您成功创建了 " & GroupName & "，请等待管理员的审核"
	else
		Sucmsg=Sucmsg+"<br>"+"<li>新门派已经成立，您成为该门派的掌门了，恭喜新掌门上任"
	end if
	Sucmsg=Sucmsg+"<br>"+"<li>您被交了 "&new_menpai_setting(6)&" 元的创派费"
	call menpai_suc()
end sub  
'--------------------------------发送短信息--------------------------------
sub sendmessage(Susername,menpainame)
	dim message,title                                        
	title=menpainame&" 清理门户通知"                                          
	message=Susername&"你好！由于你不遵守门规，已经被掌门人清理出 "&menpainame
	sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&Susername&"','"&menpainame&"掌门人','"&title&"','"&message&"',Now(),0,1)"                                          
	conn.execute(sql)                                          
end sub                                                                                                                                                                                                                            
%>