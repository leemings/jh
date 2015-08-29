<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="z_menpaiconfig.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		dim GroupID,Groupname
		select case Request("action")
			case "add"
				call  newmenpai()    ' 添加门派
			case "applynewmenpai"
				call applynewmenpai()
			case "edit"
				call edit()		'修改门派
			case "savedit"
				call savedit()	
			case "del"
				call del()		'删除门派
			case "editapply"
				call editapply()  '修改申请创建门派的条件
			case "saveditapply"
				call saveditapply()
			case "pass"			'批准新门派成立
				call pass()
			case "repass"			'重新开通已经注销的门派
				call repass()							
			case else			
				call main()
		end select 	
		conn.close
		set conn=nothing
	end if
	if founderr then call dvbbs_error()
'=========================================================	
sub main()
%>
<br>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr>
		<th height=21 align="left">门派管理 --> <a href=z_admin_menpai.asp><font color=#FFFFFF>门派列表</font></a> | <a href=z_admin_menpai.asp?action=add><font color=#FFFFFF>添加门派</font></a></th>
	</tr>
</table>		
<br>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
  <tr>
    <th height=24>门派名称</th>
    <th height=24>现任掌门</th> 
    <th height=24>创建日期</th> 
    <th height=24>性质</th> 
    <th height=24>弟子数</th> 
    <th height=24>金钱</th> 
    <th height=24>总帖数</th> 
    <th height=24>操作</th> 
  </tr>
  <%
'门派显示  开始  
	dim rs1
	dim Zangmen,TotalArticle,TotalWealth,TotalUsernum
	dim menpai_setting	
	sql="select * from [GroupName] where state<2"
	set rs=conn.execute(sql)
	if rs.eof and rs.bof  then
		response.write "<tr><td height=21 class=forumrow colspan=8><font color=red>本论坛暂时没有任何门派</font></td></tr>"
	else 
		do while not rs.eof
			TotalArticle=0
			TotalWealth=0
			TotalUsernum=0		
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
			set rs1=conn.execute("select sum(article),sum(userwealth),count(*) from [user] where usergroup='"&rs("GroupName")&"'")
			if rs1(0)<>"" then TotalArticle=rs1(0)
			if rs1(1)<>"" then TotalWealth=rs1(1)
			if rs1(2)<>"" then TotalUsernum=rs1(2)
%>
<tr>
	<td valign=middle align=center height=21 class=forumrow><a href="z_menpaiadd.asp?menu=1&groupname=<%=rs("groupname")%>" title="门派简介:<%=menpai_setting(7)%>"><%=rs("groupname")%></a></td>
	<td valign=middle align=center height=21 class=forumrow><%=Zangmen%></td>
    <td valign=middle align=center height=21 class=forumrow><%=rs("addtime")%></td>
    <td valign=middle align=center height=21 class=forumrow><%=MenpaiAttri(menpai_setting(0))%></td>
    <td valign=middle align=center height=21 class=forumrow><%=TotalUsernum%></td>
    <td valign=middle align=center height=21 class=forumrow><%=TotalWealth%></td>
    <td valign=middle align=center height=21 class=forumrow><%=TotalArticle%></td>
    <td valign=middle align=center height=21 class=forumrow><a href="z_admin_menpai.asp?id=<%=rs("id")%>&action=edit">修改</a> | <a href="z_admin_menpai.asp?id=<%=rs("id")%>&action=del" onclick="{if(confirm('您确定要删除 <%=rs("groupname")%> 吗?')){return true;}return false;}">删除</a></td>
</tr>
<%			rs.movenext
		loop
		rs.close
	end if
%>
	<tr><td height=21 class=forumRowHighlight colspan=8></td></tr>
</table>
<%	  
'门派显示  结束

'开创新门派部分  代码开始
	dim new_menpai_setting
	new_menpai_setting=split(conn.execute("select new_menpai_setting from [menpai]")(0),",")

	'set rs=conn.execute("select userwealth,usercp,userep,userpower,userclass from [user] where username='"&membername&"'")	
 	
	dim wumenwupai			'收纳不属于任何门派弟子的门派
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
<br>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr><th colspan=9 height=20>申请建立门派的条件</th></tr>
<tr>
	<td align=center height=21 class=forumRowHighlight>项目</td>
    <td align=center height=21 class=forumRowHighlight>金钱</td>
	<td align=center height=21 class=forumRowHighlight>魅力</td>
	<td align=center height=21 class=forumRowHighlight>经验</td>
	<td align=center height=21 class=forumRowHighlight>威望</td>
	<td align=center height=21 class=forumRowHighlight>文章</td>
	<td align=center height=21 class=forumRowHighlight>用户等级</td> 
	<td align=center height=21 class=forumRowHighlight>审核</td> 
	<td align=center height=21 class=forumRowHighlight>门派</td>
</tr>
<tr>
	<td align=center height=21 class=forumrow>创派要求</td>
    <td align=center height=21 class=forumrow><%=new_menpai_setting(0)%>+<font color=red><%=new_menpai_setting(6)%></font></td>
	<td align=center height=21 class=forumrow><%=new_menpai_setting(1)%></td>
	<td align=center height=21 class=forumrow><%=new_menpai_setting(2)%></td>
	<td align=center height=21 class=forumrow><%=new_menpai_setting(3)%></td>
	<td align=center height=21 class=forumrow><%=new_menpai_setting(4)%></td>
	<td align=center height=21 class=forumrow><%=GetUserTitle(new_menpai_setting(5))%>(<%=new_menpai_setting(5)%>级)</td>
	<td align=center height=21 class=forumrow><%=iif(cint(new_menpai_setting(7))=1,"需要","不需要")%></td>
	<td align=center height=21 class=forumrow><%=wumenwupai%></td>
</tr>
<tr><td align=center height=21 class=forumRowHighlight colspan=9><a href=z_admin_menpai.asp?action=editapply>修改</a></td></tr>
</table>
<%
'开创新门派部分  代码结束 

'等待审核的门派 代码开始
%>	
<br>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
  	<tr><th height=24 colspan="5">等待审核的门派</th></tr> 
  	<tr>
		<td align=center height=21 class=forumRowHighlight>门派名称</th>
		<td align=center height=21 class=forumRowHighlight>创始人</th> 
		<td align=center height=21 class=forumRowHighlight>创建日期</th> 
		<td align=center height=21 class=forumRowHighlight>性质</th> 
		<td align=center height=21 class=forumRowHighlight>操作</th> 
  	</tr>
<%
	sql="select * from [GroupName] where state=3"
	set rs=conn.execute(sql)
	if rs.eof and rs.bof  then
		response.write "<tr><td height=21 class=forumrow colspan=8><font color=navy>暂时没有等待审核的门派</font></td></tr>"
	else 
		do while not rs.eof
			menpai_setting=split(rs("menpai_setting"),"|")
			Zangmen=trim(rs("Zangmen"))
%>
<tr>
	<td valign=middle align=center height=21 class=forumrow><a href="z_menpaiadd.asp?menu=1&groupname=<%=rs("groupname")%>" title="门派简介:<%=menpai_setting(7)%>"><%=rs("groupname")%></a></td>
	<td valign=middle align=center height=21 class=forumrow><a href=dispuser.asp?name=<%=htmlencode(Zangmen)%> title="查看该掌门人的资料" target=_blank><%=Zangmen%></a></td>
    <td valign=middle align=center height=21 class=forumrow><%=rs("addtime")%></td>
    <td valign=middle align=center height=21 class=forumrow><%=MenpaiAttri(menpai_setting(0))%></td>
    <td valign=middle align=center height=21 class=forumrow><a href="z_admin_menpai.asp?id=<%=rs("id")%>&Zangmen=<%=Zangmen%>&action=pass">批准</a> | <a href="z_admin_menpai.asp?id=<%=rs("id")%>&action=edit">编辑</a> | <a href="z_admin_menpai.asp?id=<%=rs("id")%>&Zangmen=<%=Zangmen%>&action=del" onclick="{if(confirm('您确定删除 <%=rs("groupname")%> 吗?')){return true;}return false;}">删除</a></td>
</tr>
<%			rs.movenext
		loop
		rs.close
	end if
%>
	<tr><td height=21 class=forumRowHighlight colspan=5></td></tr>
</table>
<%
'等待审核的门派  代码结束

'已注销的门派  代码开始
%>	
<br>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
  	<tr><th height=24 colspan="4">已注销的门派</th></tr> 
  	<tr>
		<td align=center height=21 class=forumRowHighlight>门派名称</th>
		<td align=center height=21 class=forumRowHighlight>创建日期</th> 
		<td align=center height=21 class=forumRowHighlight>性质</th> 
		<td align=center height=21 class=forumRowHighlight>操作</th> 
  	</tr>
<%
	sql="select * from [GroupName] where state=2"
	set rs=conn.execute(sql)
	if rs.eof and rs.bof  then
		response.write "<tr><td height=21 class=forumrow colspan=8><font color=navy>暂时没有已注销的门派</font></td></tr>"
	else 
		do while not rs.eof
			menpai_setting=split(rs("menpai_setting"),"|")
%>
<tr>
	<td valign=middle align=center height=21 class=forumrow><a href="z_menpaiadd.asp?menu=1&groupname=<%=rs("groupname")%>" title="门派简介:<%=menpai_setting(7)%>"><%=rs("groupname")%></a></td>
    <td valign=middle align=center height=21 class=forumrow><%=rs("addtime")%></td>
    <td valign=middle align=center height=21 class=forumrow><%=MenpaiAttri(menpai_setting(0))%></td>
    <td valign=middle align=center height=21 class=forumrow><a href="z_admin_menpai.asp?id=<%=rs("id")%>&action=repass">重开</a> | <a href="z_admin_menpai.asp?id=<%=rs("id")%>&action=edit">编辑</a> | <a href="z_admin_menpai.asp?id=<%=rs("id")%>&action=del" onclick="{if(confirm('您确定退出 <%=rs("groupname")%> 吗?')){return true;}return false;}">删除</a></td>
</tr>
<%			rs.movenext
		loop
		rs.close
	end if
%>
	<tr><td height=21 class=forumRowHighlight colspan=5></td></tr>
</table>
<%
'已注销的门派  代码结束
end sub

'----------------------------添加新门派------------------------------
sub newmenpai()       
%>                                                                                             
<table cellspacing=1 align=center class=tableborder width="95%">                                                                                                                                                                              
	<form name="form" method="post" action="z_admin_menpai.asp?action=applynewmenpai">                                                                                                                                                                                            
  	<tr>                                                                                                                                                                                           
    	<th valign=middle align=center height=24 colspan="4">添加新门派</th>                                                                                                                                                                                          
  	</tr> 
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>                                                                                                                                                                                        
  	<tr> 
		<td valign=middle align=left class=forumRowHighlight height="22"><b>门派名称:</b><br> 门派名称申请之后掌门不能修改</td> 
		<td valign=middle align=left class=forumrow height="22" colspan="3">&nbsp;<INPUT name=groupname size="20" maxlength="20"> <font color=red>*</font> (不要超过12个英文字母或6个汉字)</td>
  	</tr>
  	<tr> 
		<td valign=middle align=left class=forumRowHighlight height="22"><b>掌门:</b><br> 如果暂时不想设置掌门，可以先留空</td> 
		<td valign=middle align=left class=forumrow height="22" colspan="3">&nbsp;<INPUT name=Zangmen size="20" maxlength="20" value=<%=membername%>></td>
  	</tr>
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>
    <tr>
		<td valign=middle align=left class=forumRowHighlight height="22"><b>门派性质说明:</b></td>                                                                                                                                                                                         
		<td valign=middle align=left class=forumrow height="22" colspan="3">
			<li>中立派是收留所有无门无派的人的一个特殊门派
			<li>新注册用户和被掌门踢出门派的用户自动加入中立派
		 	<li>只能存在一个中立派，所以当您设置一个门派为中立派之后，以前的中立派自动转为正常门派
		</td> 
  	</tr>		
  	<tr> 
		<td valign=middle align=left class=forumRowHighlight height="22"><b>门派性质:</b><br> 门派性质申请之后掌门不能修改</td>
		<td valign=middle align=left class=forumrow height="22" colspan="3">
			<input type=radio name=menpaiattri value=0 checked>明门正派&nbsp;
			<input type=radio name=menpaiattri value=1>社团&nbsp;
			<input type=radio name=menpaiattri value=2>魔教&nbsp;
			<input type=radio name=menpaiattri value=3>中立派<br>
		</td>
  	</tr>
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>
    <tr>
		<td valign=middle align=left class=forumRowHighlight height="22"><b>门限值的说明:</b></td>                                                                                                                                                                                         
		<td valign=middle align=left class=forumrow height="22" colspan="3">
		 	<li>金钱,魅力,经验,威望,文章数,用户等级 都不能超过您自已的当前值
			<li>金钱,魅力,经验,威望,文章数,用户等级 是加入本派的必须满足的最小值
			<li>金钱,魅力,经验,威望,文章数,用户等级 也是加入和退出门派增加/扣除相关值时的基数
			<li>明门正派不受以上限制
			<li>以下属性掌门可以在门派管理中修改			
		</td> 
  	</tr> 
    <tr>                                                                                                                                                                                         
		<td valign=middle align=left class=forumRowHighlight height="22" width="25%"><b>金钱:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userwealth value=""></td>                                                                                                                                                                                        
		<td valign=middle align=right class=forumRowHighlight height="22">您的金钱:&nbsp;</td>
		<td valign=middle align=left class=forumrow height="22" width="20%">&nbsp;<%=mymoney%></td>
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>经验:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userep value=""></td>
		<td valign=middle align=right class=forumRowHighlight height="22">您的经验:&nbsp;</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myusercp%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>魅力:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=usercp value=""></td> 
		<td valign=middle align=right class=forumRowHighlight height="22">您的魅力:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myuserep%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>威望:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userpower value=""></td> 
		<td valign=middle align=right class=forumRowHighlight height="22">您的威望:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=mypower%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>文章数:</b><br>加入本派必须达到的发帖数</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userarticle value=""></td> 
		<td valign=middle align=right class=forumRowHighlight height="22">您的帖数:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myarticle%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>用户等级:</b><br>加入本派必须具备的等级(1～<%=maxtitleid%>)</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userclass value=""></td> 
		<td valign=middle align=right class=forumRowHighlight height="22">您的等级:&nbsp;</td>                                                                                                                                                                                     
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myclass%>（<%=GetUserClass(myclass)%>级）</td>                                                                                                                                                                                      
  	</tr>
	<tr>
		<td valign=middle align=center class=forumrow height=22 colspan="4"><font color="#CC9933">
		如果门派性质是<font color=red>明门正派</font>或<font color=red>中立派</font>，则以上6项(金钱,魅力,经验,威望,文章数,用户等级)不必填写</font>
		</td>
	</tr> 
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>                                                                                                                                                                                       
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>本派简介:</b><br>对您的门派作简单的介绍，在您的门派资料中会显示出来</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow colspan="3" height="22">&nbsp;<textarea name=menpaireadme rows="4" cols="60"></textarea>		(不要超过100个英文或50个汉字)</td>                                                                                                                                                                                      
  	</tr>  
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>                                                                                                                                                                                       
  	<tr>                                                                                                                                                                                       
    	<td valign=middle align=center class=forumrow colspan="4" height="22"><INPUT type=submit value=添加 name=submit></td>                                                                                                                                                                                       
  	</tr>                                                                                                                                                                              
  	</form>                                                                                                                                                                                    
</table>                                                                                                                                                                                           
<%                                             
end sub                                                                                                                                                                                                                            

'------------------------------------添加门派成功--------------------------------
sub applynewmenpai() 
	
	dim MenpaiAttri,userwealth,usercp,userep,userpower,userclass,userartitle,Zangmen
	GroupName=checkstr(trim(Request.form("groupname")))
	MenpaiAttri=Request.form("menpaiattri")
	Zangmen=trim(Request.form("Zangmen"))
	
	if GroupName="" then
		Errmsg=Errmsg+"<br>"+"<li>请输入门派名称"
		founderr=true
	elseif strLength(GroupName)>12 then
		Errmsg=Errmsg+"<br>"+"<li>门派名称不能超过12个英文字母(6个汉字)"
		founderr=true		
	end if		
	if cint(MenpaiAttri)<>0 and cint(MenpaiAttri)<>3 then
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
			Errmsg=Errmsg+"<br>"+"<li>非明门正派必须填写等级或者您填入的不是有效的数值"
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
	
	if Zangmen<>"" then
		set rs=conn.execute("select usergroup from [user] where username='"&Zangmen&"'")
		if rs.eof and rs.bof then
			Errmsg=Errmsg+"<br>"+"<li>您指定为掌门的用户不存在"
			founderr=true
		end if
	end if
				
	if founderr then exit sub
                           
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
		rs("Zangmen")=Zangmen
		rs("Addtime")=date()
		rs("Menpai_setting")=MenpaiAttri & "|" & userwealth & "|" & userep & "|" & usercp & "|" & userpower & "|" & userartitle & "|" & userclass & "|" & menpaireadme
		if cint(MenpaiAttri)=3 then
			rs("State")=1
			conn.execute("update [GroupName] set State=0 where State=1")
		else
			rs("State")=0
		end if	
		rs.update  
		if Zangmen<>"" then conn.execute("update [user] set usergroup='"&GroupName&"' where username='"&Zangmen&"'")		
	end if
%>
<br><p><b>添加门派成功</b>&nbsp;&nbsp;<a href=z_admin_menpai.asp>返回</a></p>
<%	 
end sub 
 
sub edit() 
	GroupID=request("ID")
	if GroupID="" and (not isnumeric(GroupID)) then
		Errmsg=Errmsg+"<br>"+"<li>错误的门派参数"
		founderr=true
		exit sub		
	end if
	dim Zangmen,Addtime,Menpai_setting
	sql="select * from [GroupName] where id="&GroupID&""
	set rs=conn.execute(sql)
	if not(rs.eof and rs.bof) then
		GroupName=rs("GroupName")
		Zangmen=trim(rs("Zangmen"))
		Addtime=rs("Addtime")
		Menpai_setting=split(rs("Menpai_setting"),"|")
	else
		Errmsg=Errmsg+"<br>"+"<li>错误的门派参数,指定门派不存在"
		founderr=true
		exit sub
	end if	
	rs.close
%>   
<table cellspacing=1 align=center class=tableborder width="95%">                                                                                                                                                                              
  	<tr>                                                                                                                                                                                           
    	<td valign=middle align=center height=24 class=forumRowHighlight colspan="4">门派管理</td>                                                                                                                                                                                          
  	</tr> 
	<form name="form" method="post" action="z_admin_menpai.asp?action=savedit">                                                                                                                                                                                            
  	<tr>                                                                                                                                                                                           
    	<th valign=middle align=left height=24 colspan="4">&nbsp;★&nbsp; 门派基本设置</th>                                                                                                                                                                                          
  	</tr> 
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>                                                                                                                                                                                        
  	<tr> 
		<td valign=middle align=left class=forumRowHighlight height="22"><b>门派名称:</b><br> 门派名称掌门不能修改</td> 
		<td valign=middle align=left class=forumrow height="22" colspan="3">&nbsp;<INPUT name=GroupName size="20" maxlength="20" value=<%=GroupName%>> <font color=red>*</font> (不要超过12个英文字母或6个汉字)<INPUT name=GroupID size="20" maxlength="20" type=hidden value=<%=GroupID%>></td>
  	</tr>
  	<tr> 
		<td valign=middle align=left class=forumRowHighlight height="22"><b>掌门:</b><br> 如果不想设置任何人做掌门，可以留空</td> 
		<td valign=middle align=left class=forumrow height="22" colspan="3">&nbsp;<INPUT name=Zangmen size="20" maxlength="20" value=<%=Zangmen%>></td>
  	</tr>
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>
    <tr>
		<td valign=middle align=left class=forumRowHighlight height="22"><b>门派性质说明:</b></td>                                                                                                                                                                                         
		<td valign=middle align=left class=forumrow height="22" colspan="3">
			<li>中立派是收留所有无门无派的人的一个特殊门派
			<li>新注册用户和被掌门踢出门派的用户自动加入中立派		
		 	<li>只能存在一个中立派，所以当您设置一个门派为中立派之后，以前的中立派自动转为正常门派
			<li>选择明门正派或中立派会清除以前设置的金钱,魅力,经验,威望,文章数,用户等级的值(都为0)
		</td> 
  	</tr>		
  	<tr> 
		<td valign=middle align=left class=forumRowHighlight height="22"><b>门派性质:</b><br> 门派性质申请之后掌门不能修改</td>
		<td valign=middle align=left class=forumrow height="22" colspan="3">
			<input type=radio name=MenpaiAttri value=0 <%if cint(menpai_setting(0))=0 then%>checked<%end if%>>明门正派&nbsp;
			<input type=radio name=MenpaiAttri value=1 <%if cint(menpai_setting(0))=1 then%>checked<%end if%>>社团&nbsp;
			<input type=radio name=MenpaiAttri value=2 <%if cint(menpai_setting(0))=2 then%>checked<%end if%>>魔教&nbsp;
			<input type=radio name=MenpaiAttri value=3 <%if cint(menpai_setting(0))=3 then%>checked<%end if%>>中立派
		</td>
  	</tr>
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>
  	<tr> 
		<td valign=middle align=left class=forumRowHighlight height="22"><b>建立时间:</b><br> 门派创建的时间，格式 2002-9-21 </td> 
		<td valign=middle align=left class=forumrow height="22" colspan="3">&nbsp;<INPUT name=Addtime size="20" maxlength="20" value=<%=Addtime%>></td>
  	</tr>	
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>
    <tr>
		<td valign=middle align=left class=forumRowHighlight height="22"><b>门限值说明:</b></td>                                                                                                                                                                                         
		<td valign=middle align=left class=forumrow height="22" colspan="3">
		 	<li>金钱,魅力,经验,威望,文章数,用户等级 都不能超过您自已的当前值
			<li>金钱,魅力,经验,威望,文章数,用户等级 是加入本派的必须满足的最小值
			<li>金钱,魅力,经验,威望,文章数,用户等级 也是加入和退出门派增加/扣除相关值时的基数
			<li>明门正派和中立派 不受以上限制
		</td> 
  	</tr> 
    <tr>                                                                                                                                                                                         
		<td valign=middle align=left class=forumRowHighlight height="22" width="30%"><b>金钱:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userwealth value=<%=menpai_setting(1)%>></td>                                                                                                                                                                                        
		<td valign=middle align=right class=forumRowHighlight height="22">您的金钱:&nbsp;</td>
		<td valign=middle align=left class=forumrow height="22" width="20%">&nbsp;<%=mymoney%></td>
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>经验:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userep value=<%=menpai_setting(2)%>></td>
		<td valign=middle align=right class=forumRowHighlight height="22">您的经验:&nbsp;</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myusercp%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>魅力:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=usercp value=<%=menpai_setting(3)%>></td> 
		<td valign=middle align=right class=forumRowHighlight height="22">您的魅力:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myuserep%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>威望:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userpower value=<%=menpai_setting(4)%>></td> 
		<td valign=middle align=right class=forumRowHighlight height="22">您的威望:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=mypower%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>文章数:</b><br>加入本派必须达到的发帖数</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userarticle value=<%=menpai_setting(5)%>></td> 
		<td valign=middle align=right class=forumRowHighlight height="22">您的帖数:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myarticle%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>用户等级:</b><br>加入本派必须具备的等级(1～<%=maxtitleid%>)</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userclass value=<%=menpai_setting(6)%>>&nbsp;<%=GetUserTitle(Menpai_setting(6))%>(<%=Menpai_setting(6)%>级)</td> 
		<td valign=middle align=right class=forumRowHighlight height="22">您的等级:&nbsp;</td>                                                                                                                                                                                     
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myclass%>（<%=GetUserClass(myclass)%>级）</td>                                                                                                                                                                                      
  	</tr>
	<tr>
		<td valign=middle align=center class=forumrow height=22 colspan="4"><font color="#CC9933">
		如果门派性质是<font color=red>明门正派</font>或<font color=red>中立派</font>，则以上6项(金钱,魅力,经验,威望,文章数,用户等级)不必填写</font>
		</td>
	</tr> 
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>                                                                                                                                                                                       
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>本派简介:</b><br>对您的门派作简单的介绍，在您的门派资料中会显示出来</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow colspan="3" height="22">&nbsp;<textarea name=menpaireadme rows="4" cols="60"><%=menpai_setting(7)%></textarea>		(不要超过100个英文或50个汉字)</td>                                                                                                                                                                                      
  	</tr>  
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>                                                                                                                                                                                       
  	<tr>                                                                                                                                                                                       
    	<td valign=middle align=center class=forumrow colspan="4" height="22"><INPUT type=submit value=更新 name=submit></td>                                                                                                                                                                                       
  	</tr>                                                                                                                                                                              
  	</form>     
</table>    
                                                                                                                                                                           
<%                                                                                                                                                                                                                             
end sub

sub savedit()
	                                                                                                                                                   
	dim MenpaiAttri,userwealth,usercp,userep,userpower,userclass,userartitle,Addtime,Zangmen
	GroupID=Request.form("GroupID")
	GroupName=checkstr(trim(Request.form("GroupName")))
	MenpaiAttri=Request.form("MenpaiAttri")
	Zangmen=trim(Request.form("Zangmen"))
	Addtime=Request.form("Addtime")
	
	if GroupName="" then
		Errmsg=Errmsg+"<br>"+"<li>请输入门派名称"
		founderr=true
	elseif strLength(GroupName)>12 then
		Errmsg=Errmsg+"<br>"+"<li>门派名称不能超过12个英文字母(6个汉字)"
		founderr=true		
	end if		
	if cint(MenpaiAttri)<>0 and cint(MenpaiAttri)<>3 then
		if not isInteger(Request.form("userwealth")) then
			Errmsg=Errmsg+"<br>"+"<li>非明门正派必须填写金钱值或者您填入的不是有效的数值"
			founderr=true
		else
			userwealth=int(Request.form("userwealth"))
		end if	
		if not isInteger(Request.form("userep")) then
			Errmsg=Errmsg+"<br>"+"<li>非明门正派必须填写经验值或者您填入的不是有效的数值"
			founderr=true
		else 
			userep=int(Request.form("userep"))	
		end if
		if not isInteger(Request.form("usercp")) then
			Errmsg=Errmsg+"<br>"+"<li>非明门正派必须填写魅力值或者您填入的不是有效的数值"
			founderr=true
		else 
			usercp=int(Request.form("usercp"))	
		end if
		if not isInteger(Request.form("userpower")) then
			Errmsg=Errmsg+"<br>"+"<li>非明门正派必须填写威望值或者您填入的不是有效的数值"
			founderr=true
		else 
			userpower=int(Request.form("userpower"))	
		end if
		if not isInteger(Request.form("userarticle")) then
			Errmsg=Errmsg+"<br>"+"<li>非明门正派必须填写文章数或者您填入的不是有效的数值"
			founderr=true
		else 
			userartitle=int(Request.form("userarticle"))	
		end if
		if not isInteger(Request.form("userclass")) then
			Errmsg=Errmsg+"<br>"+"<li>非明门正派必须填写威望值或者您填入的不是有效的数值"
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
	
	if Zangmen<>"" then
		set rs=conn.execute("select usergroup from [user] where username='"&Zangmen&"'")
		if rs.eof and rs.bof then
			Errmsg=Errmsg+"<br>"+"<li>您指定为掌门的用户不存在"
			founderr=true
		end if
	end if				 	
	
	dim menpaireadme
	menpaireadme=Request.form("menpaireadme")	
	if strLength(menpaireadme)>100 then
		Errmsg=Errmsg+"<br>"+"<li>门派简介不能超过100个英文或50个汉字"
		founderr=true		
	end if
			
	if founderr then exit sub
	
	set rs= server.createobject ("adodb.recordset")                                                                                                                                                                                                                                   
	sql="select * from [GroupName] where ID="&GroupID&""                                                                                                                                                                                                                                   
	rs.Open sql,conn,1,3  
	 
	if rs.bof and rs.eof then                                                                          
		Errmsg=Errmsg+"<br>"+"<li>该门派不存在啊"
		founderr=true
		rs.close
		set rs=nothing
		exit sub
	else                                  
		rs("Groupname")=GroupName
		rs("Zangmen")=Zangmen
		if cint(MenpaiAttri)=3 then
			rs("State")=1
			conn.execute("update [GroupName] set State=0 where State=1")
		else
			rs("State")=0
		end if		
		rs("Menpai_setting")=MenpaiAttri & "|" & userwealth & "|" & userep & "|" & usercp & "|" & userpower & "|" & userartitle & "|" & userclass & "|" & menpaireadme
		if isdate(Addtime) then rs("Addtime")=Addtime
		rs.update  
		if Zangmen<>"" then conn.execute("update [user] set usergroup='"&GroupName&"' where username='"&Zangmen&"'")
	end if
%>
<br><br><p><b>修改门派成功</b>&nbsp;&nbsp;<a href=z_admin_menpai.asp>返回</a></p>
<%
end sub

sub editapply() 
	dim newmenpai_setting
	newmenpai_setting=split(conn.execute("select new_menpai_setting from [menpai]")(0),",")
%>   
<br>
<table cellspacing=1 align=center class=tableborder width="95%">                                                                                                                                                                              
  	<tr>                                                                                                                                                                                           
    	<th valign=middle height=24 colspan="4">门派申请条件管理</td>                                                                                                                                                                                          
  	</tr> 
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>        
	<form name="form" method="post" action="z_admin_menpai.asp?action=saveditapply">                                                                                                                                                                                 
    <tr>
		<td valign=middle align=left class=forumRowHighlight height="22"><b>说明:</b></td>                                                                                                                                                                                         
		<td valign=middle align=left class=forumrow height="22" colspan="3">
			<li>前台用户可以自己创建门派必须满足的条件
			<li>条件最好不要设置的太低，否则很多人都符号创派要求
			<li>您的属性值是用来参考的，输入的值不受您的属性值的限制
			<li>请输入数字,不要输入负数和字母		
		</td> 
  	</tr> 
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>
    <tr>                                                                                                                                                                                         
		<td valign=middle align=left class=forumRowHighlight height="22" width="30%"><b>金钱:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userwealth value=<%=newmenpai_setting(0)%>></td>                                                                                                                                                                                        
		<td valign=middle align=right class=forumRowHighlight height="22">您的金钱:&nbsp;</td>
		<td valign=middle align=left class=forumrow height="22" width="20%">&nbsp;<%=mymoney%></td>
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>经验:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userep value=<%=newmenpai_setting(2)%>></td>
		<td valign=middle align=right class=forumRowHighlight height="22">您的经验:&nbsp;</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myusercp%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>魅力:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=usercp value=<%=newmenpai_setting(1)%>></td> 
		<td valign=middle align=right class=forumRowHighlight height="22">您的魅力:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myuserep%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>威望:</b></td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userpower value=<%=newmenpai_setting(3)%>></td> 
		<td valign=middle align=right class=forumRowHighlight height="22">您的威望:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=mypower%></td>                                                                                                                                                                                        
  	</tr>                                                                                                                                                                                        
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>文章数:</b><br>加入本派必须达到的发帖数</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userarticle value=<%=newmenpai_setting(4)%>></td> 
		<td valign=middle align=right class=forumRowHighlight height="22">您的帖数:&nbsp;</td>                                                                                                                                                                                       
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myarticle%></td>                                                                                                                                                                                        
  	</tr> 
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>用户等级:</b><br>加入本派必须具备的等级(1～<%=maxtitleid%>)</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22">&nbsp;<INPUT name=userclass value=<%=newmenpai_setting(5)%>>&nbsp;<%=GetUserTitle(newmenpai_setting(5))%>(<%=newmenpai_setting(5)%>级)</td> 
		<td valign=middle align=right class=forumRowHighlight height="22">您的等级:&nbsp;</td>                                                                                                                                                                                     
		<td valign=middle align=left class=forumrow height="22">&nbsp;<%=myclass%>（<%=GetUserClass(myclass)%>级）</td>                                                                                                                                                                                      
  	</tr>
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>  
  	<tr>                                                                                                                                                                                        
		<td valign=middle align=left class=forumRowHighlight height="22"><b>创派收费:</b><br>申请门派时收取的费用</td>                                                                                                                                                                                        
		<td valign=middle align=left class=forumrow height="22" colspan="3">&nbsp;<INPUT name=creatmoney value=<%=newmenpai_setting(6)%>>
		&nbsp;<font color=navy>注意：前台创派条件中的金钱项 = 前面的金钱 + <font color=red>创派收费</font></font>
		</td> 
  	</tr>
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>  
  	<tr>                                                                                                                                                                                         
		<td valign=middle align=left class=forumRowHighlight height="22"><b>审核选项:</b><br>用户申请门派之后是否需要通过管理员审核之后才能正式成立</td>                                                                                                                                                                                         
		<td valign=middle align=left class=forumrow height="22" colspan="3">&nbsp;
			<input type=radio name=shenhe value=0 <%if cint(newmenpai_setting(7))=0 then%>checked<%end if%>>直接开通&nbsp;
			<input type=radio name=shenhe value=1 <%if cint(newmenpai_setting(7))=1 then%>checked<%end if%>>需要通过审核后才能开通&nbsp;
		</td> 
  	</tr>
	<tr><td height=8 class=forumRowHighlight colspan="4"></td></tr>  		                                                                                                                                                                                     
  	<tr>                                                                                                                                                                                       
    	<td valign=middle align=center class=forumrow colspan="4" height="22"><INPUT type=submit value=更新 name=submit></td>                                                                                                                                                                                       
  	</tr>                                                                                                                                                                              
  	</form>     
</table>    
<%                                                                                                                                                                                                                             
end sub

sub saveditapply()
	dim maxtitleid,creatmoney
	dim MenpaiAttri,userwealth,usercp,userep,userpower,userclass,userartitle,Addtime,Zangmen

	if not isInteger(Request.form("userwealth")) then
		Errmsg=Errmsg+"<br>"+"<li>您输入的金钱值不是有效的数值"
		founderr=true
	else
		userwealth=int(Request.form("userwealth"))
	end if	
	if not isInteger(Request.form("userep")) then
		Errmsg=Errmsg+"<br>"+"<li>您输入的经验值不是有效的数值"
		founderr=true
	else 
		userep=int(Request.form("userep"))	
	end if
	if not isInteger(Request.form("usercp")) then
		Errmsg=Errmsg+"<br>"+"<li>您输入的魅力值不是有效的数值"
		founderr=true
	else 
		usercp=int(Request.form("usercp"))	
	end if
	if not isInteger(Request.form("userpower")) then
		Errmsg=Errmsg+"<br>"+"<li>您输入的威望值不是有效的数值"
		founderr=true
	else 
		userpower=int(Request.form("userpower"))	
	end if
	if not isInteger(Request.form("userarticle")) then
		Errmsg=Errmsg+"<br>"+"<li>您输入的文章值不是有效的数值"
		founderr=true
	else 
		userartitle=int(Request.form("userarticle"))	
	end if
	maxtitleid=conn.execute("select top 1 UserTitleID from [UserTitle] order by UserTitleID desc")(0)
	if not isInteger(Request.form("userclass")) then
		Errmsg=Errmsg+"<br>"+"<li>您输入的等级值不是有效的数值"
		founderr=true
	elseif int(Request.form("userclass"))<1 or int(Request.form("userclass"))>int(maxtitleid) then
		Errmsg=Errmsg+"<br>"+"<li>等级值不能小于1或大于"&maxtitleid
		founderr=true						
	else 
		userclass=int(Request.form("userclass"))	
	end if
	if not isInteger(Request.form("creatmoney")) then
		Errmsg=Errmsg+"<br>"+"<li>您输入的创派费用不是有效的数值"
		founderr=true
	else
		creatmoney=int(Request.form("creatmoney"))
	end if	
	
	if founderr then exit sub
	conn.execute("update menpai set new_menpai_setting='" & userwealth & "," & usercp & "," & userep & "," & userpower & "," & userartitle & "," & userclass & "," & creatmoney & "," & Request.form("shenhe") & "'")
%>
<br><br><p>&nbsp;<b>修改创派条件成功</b>&nbsp;&nbsp;<a href=z_admin_menpai.asp>返回</a></p>
<%	
end sub
'------------------删除门派----------------------
sub del()
	conn.execute("delete from GroupName where id="&request("id"))
	if request("Zangmen")<>"" then call sendmessage(request("Zangmen"),membername,"门派审核通知","真是抱歉，您新创建的门派没有通过审核")
%>
<br><br><p><b>删除门派成功</b>&nbsp;&nbsp;<a href=z_admin_menpai.asp>返回</a></p>
<%
end sub
'------------------批准门派成立------------------
sub pass()
	conn.execute("update GroupName set state=0 where id="&request("id"))
	if request("Zangmen")<>"" then	call sendmessage(request("Zangmen"),membername,"门派开通通知","经过审核，您开创的门派已经可以开通了，希望您能够管理好您的门派，谢谢!")
%>
<br><br><p><b>您批准了该门派成立</b>&nbsp;&nbsp;<a href=z_admin_menpai.asp>返回</a></p>
<%	
end sub 
'------------------重新开通已经注销的门派------------------
sub repass()
	conn.execute("update GroupName set state=0,zangmen='' where id="&request("id"))
%>
<br><br><p><b>门派重开成功，请在门派管理中修改门派的基本资料和掌门</b>&nbsp;&nbsp;<a href=z_admin_menpai.asp>返回</a></p>
<%	
end sub 
'--------------------------------发送短信息--------------------------------
sub sendmessage(incept,sender,title,message)
	sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&incept&"','"&sender&"','"&title&"','"&message&"',Now(),0,1)"                                          
	conn.execute(sql)                                          
end sub 
%>