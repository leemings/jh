<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_plus_conn.asp"-->
<%response.buffer=true%>
<head>
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	call dvbbs_error()
else
	dim body,Plus_Config
	dim readme,Tlink
	call main()
	set rs=nothing
end if

sub main()
	select case request("action")
	case "tj"
		call tj()
	case "plusinfo"
		call plusinfo()
	case "add"
		call addlink()
	case "savenew"
		call savenew()
	case "edit"
		call editlink()
	case "savedit"
		call savedit()
	case "del"
		call del()
	case "OrdersUp"
		call OrdersUp()
	case "OrdersDown"
		call OrdersDown()
	case "addGroup"
		call addGroup()
	case "Groupedit"
		call Groupedit()
	case "Groupdel"
		call Groupdel()
	case "GroupOrdersUp"
		call GroupOrdersUp()
	case "GroupOrdersDown"
		call GroupOrdersDown()
	case else
		call Groupinfo()
	end select
	response.write body
end sub

sub Groupinfo()
	dim rsGroup
	set rs= server.createobject ("adodb.recordset")
	sql = "select * from [Group] order by GroupID"
	rs.open sql,connplus,1,1%>
	<table width="100%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	  <tr> 
	    <th width="18%" height=25 class="tableHeaderText">组ID</th>
	    <th width="25%" class="tableHeaderText">组内插件数</th>
	    <th width="28%" class="tableHeaderText">组名（点击查看组内插件）</th>
	    <th width="29%" class="tableHeaderText">操作</th>
	  </tr>
		<%if rs.eof and rs.bof then%>
		  <tr> 
		    <td height=25 colspan="4" class=forumrow align=center>没有设定插件组！</td>
		  </tr>
		<%else%>
			<%do while not rs.eof%>
			  <tr> 
			    <td height=25 class=forumrow align=center><%=rs("Groupid")%></td>
			    <td class=forumrow align=center><%
			    	set rsGroup= server.createobject ("adodb.recordset")
						rsGroup.open ("select * from plus where Groupid=" & rs("Groupid"))&" order by groupid",connplus,1,1
						response.write rsGroup.recordcount
						rsGroup.close
					%></td>
			  	<td class=forumrow align=center><a href="?action=plusinfo&Groupid=<%=rs("Groupid")%>&Groupname=<%=rs("Groupname")%>"><%=rs("Groupname")%></a></td>
			  	<td class=forumrow align=center><a href="?action=add&Groupid=<%=rs("Groupid")%>&Groupname=<%=rs("Groupname")%>">增加</a> <a href="?action=Groupedit&groupid=<%=rs("Groupid")%>">编辑</a> <a href="?action=Groupdel&Groupid=<%=rs("Groupid")%>">删除</a> <a href="?action=GroupOrdersUp&Groupid=<%=rs("Groupid")%>">向上</a> <a href="?action=GroupOrdersDown&Groupid=<%=rs("Groupid")%>">向下</a></td>
			  </tr>
			  <%rs.movenext
			loop
			set rsGroup=nothing
		end if%>
	  <tr> 
	    <td colspan=4 class=forumrow align=right height=25>[<a href="?action=addGroup"><b>增加新的插件组</b></a>]&nbsp;&nbsp;&nbsp;&nbsp;</td>
	  </tr>
	</table>
	<p align="center"><font color="#FF0000">注意：删除一个组，组里的全部插件也会一起删除！</font></p>
	<%rs.close
	set rs=nothing
end sub

sub addGroup
	if not isnull(request("add")) and request("add")<>"" then
		if request("Groupname")="" then body=body+"<br>"+"组名不能为空。":exit sub
		dim Groupid
		set rs=connplus.execute ("select * from [Group] where Groupname='"&request("Groupname")&"'")
		if not (rs.bof and rs.eof) then body=body + "<br>组名不能重复！请重新输入！":exit sub
		set rs= server.createobject ("adodb.recordset")
		sql = "select * from [Group] order by Groupid desc"
		rs.open sql,connplus,1,3
		if rs.bof and rs.eof then Groupid=1 else Groupid=rs("Groupid")+1
		rs.addnew
		rs("Groupid")=Groupid
		rs("Groupname")=request("Groupname")
		rs.update
		rs.close
		set rs=nothing
		response.redirect "?action=Groupinfo"
	end if%>
  <form action="?action=addGroup&add=1" method = post>
	<table width="100%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	  <tr> 
	    <th height=25 colspan="2" class="tableHeaderText"><a href="?action=Groupinfo"><b>插件组管理</b></a> | 增加新的组</th>
	  </tr>
		<tr> 
	    <td width="50%" height=25 class=forumrow align=right>组名&nbsp;</td>
	    <td width="50%" class=forumrow align=left>&nbsp;<input name="Groupname" type="text" size=40></td>
		</tr>
		<tr> 
			<td height=25 colspan="2" class=forumrow align=center><input type="submit" name="Submit" value="添 加"></td>
		</tr>
	</table>
	</form>
<%end sub

sub Groupedit()
	if not isnull(request("edit")) and request("edit")<>"" then
		if request("Groupname")="" then body=body + "<br>"+"组名字不能为空！请重新输入！":exit sub
		connplus.execute "update [Group] set Groupname='" & request("Groupname") &"' where groupid="&request("groupid")
		connplus.execute "update Plus set Groupname='" & request("Groupname") &"' where groupid="&request("groupid")
		response.redirect "?action=Groupinfo"
	end if
	set rs= server.createobject ("adodb.recordset")
	sql = "select * from [Group] where groupid="&request("groupid")
	rs.open sql,connplus,1,1%>
  <form action="?action=Groupedit&edit=1" method=post>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<input type=hidden name=groupid value=<%=Request("Groupid")%>>
    <tr> 
      <th height=25 colspan="2" class="tableHeaderText"><a href="?action=Groupinfo"><b>插件组管理</b></a> | <a href="?action=addGroup"><b>增加新的插件组</b></a></th>
    </tr>
    <tr> 
      <td height=25 width="50%" class=forumrow align=right>组ID&nbsp;</td>
      <td width="50%" class=forumrow align=left>&nbsp;<%=rs("Groupid")%></td>
    </tr>
    <tr> 
      <td height=25 width="50%" class=forumrow align=right>组名称&nbsp;</td>
      <td width="50%" class=forumrow align=left>&nbsp;<input name="Groupname" type="text" value=<%=rs("Groupname")%> size=40></td>
    </tr>
    <tr> 
      <td height=25 colspan="2" class=forumrow align=center><input type="submit" name="Submit" value="修 改"></td>
    </tr>
	</table>
  </form>
	<%rs.close
	set rs=nothing
end sub

sub Groupdel()
	connplus.execute "delete * from [Group] where Groupid="& request("Groupid")
	connplus.execute "delete * from plus where Groupid="& request("Groupid")
	response.redirect "?action=Groupinfo"
end sub

sub GroupOrdersUp()
	dim oldid,newid
	set rs= server.createobject ("adodb.recordset")
	sql="select top 2 groupid from [group] where groupid<="&request("groupid")&" order by groupid desc"
	rs.open sql,connplus,1,3
	if not (rs.eof or rs.bof) then
		oldid=rs("groupid")
		rs.movenext
		if not (rs.eof or rs.bof) then
			newid=rs("groupid")
			rs("groupid")=oldid
			rs.update
			rs.MovePrevious
			rs("groupid")=newid
			rs.update
		end if
	end if
	rs.close
	set rs=nothing
	response.redirect "?action=Groupinfo"
end sub

sub GroupOrdersDown()
	dim oldid,newid
	set rs= server.createobject ("adodb.recordset")
	sql="select top 2 groupid from [group] where groupid>="&request("groupid")&" order by groupid"
	rs.open sql,connplus,1,3
	if not (rs.eof or rs.bof) then
		oldid=rs("groupid")
		rs.movenext
		if not (rs.eof or rs.bof) then
			newid=rs("groupid")
			rs("groupid")=oldid
			rs.update
			rs.MovePrevious
			rs("groupid")=newid
			rs.update
		end if
	end if
	rs.close
	set rs=nothing
	response.redirect "?action=Groupinfo"
end sub

sub plusinfo()
	set rs= server.createobject ("adodb.recordset")
	sql = " select * from plus where Groupid="&request("Groupid")&" order by plusid"
	rs.open sql,connplus,1,1%> 
	<table width="100%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	  <tr> 
	    <th width="85" height=25 class="tableHeaderText">插件ID</th>
	    <th width="144" class="tableHeaderText">插件名</th>
	    <th width="144" class="tableHeaderText">启用标识</th>
	    <th width="132" class="tableHeaderText">是否在新页面打开</th>
	    <th width="*" class="tableHeaderText">操作</th>
	  </tr>
	  <%if not rs.eof then
			do while not rs.eof
				Tlink=split(rs(2),"$")%>
			  <tr align=center> 
			    <td height=25 class=forumrow width="85"><font color=<%=forum_body(8)%>><%=rs("plusid")%></font></td>
			    <td class=forumrow width="144"><a href="?action=edit&plusid=<%=rs("plusid")%>&Groupid=<%=request("Groupid")%>&Groupname=<%=request("Groupname")%>"><%=rs("plusname")%></a></td>
	    		<%Plus_Config=split(rs("config"),",")
		      if Plus_Config(0)="1" then
		         Plus_Config(2)="禁用"
		      else
		         Plus_Config(2)="启用"
		      end if%>
	    		<td class=forumrow width="144"><%=Plus_Config(2)%></a></td>
			    <%dim window 
					if rs("window")="target=_blank" then
						window="是"
					elseif rs("window")="limit" then
						window="是(限制)"
					else
						window="否"
					end if%>
	    		<td class=forumrow width="132"><%=window%></td>
	    		<td class=forumrow width="*"><a href="?action=edit&plusid=<%=rs("plusid")%>&Groupid=<%=request("Groupid")%>&Groupname=<%=request("Groupname")%>">编辑</a> <a href="?action=del&plusid=<%=rs("plusid")%>&Groupid=<%=request("Groupid")%>&Groupname=<%=request("Groupname")%>">删除</a> <%if Plus_Config(4)="1" then%><a href=<%=Plus_Config(5)%>>管理</a><%else%>　　<%end if%> <a href="?action=OrdersUp&plusid=<%=rs("plusid")%>&Groupid=<%=request("Groupid")%>&Groupname=<%=request("Groupname")%>">向上</a> <a href="?action=OrdersDown&plusid=<%=rs("plusid")%>&Groupid=<%=request("Groupid")%>&Groupname=<%=request("Groupname")%>">向下</a></td>
	  		</tr>
	  		<%rs.movenext
			loop
		else
			response.write "<tr><td height=25 colspan='5' class=forumrow align=center>目前这个组中还没有插件！</td></tr>"
		end if
		rs.Close
		set rs=nothing%>
	  <tr> 
	    <td height="25" colspan=2 class=forumrow align=left nowrap>&nbsp;&nbsp;<font color=<%=forum_body(8)%>><%=request("Groupname")%>->插件列表</b</td>
	    <td colspan=3 class=forumrow align=right>[<a href="?action=Groupinfo"><b>插件组列表</b></a>] | [<a href="?action=add&Groupid=<%=request("Groupid")%>&Groupname=<%=request("Groupname")%>"><b>增加新的插件</b></a>]&nbsp;&nbsp;&nbsp;&nbsp;</th>
	  </tr>
	</table>
<%end sub

sub addlink()%>
	<form action="?action=savenew" method = post>
	<table width="100%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<input type="hidden" name="Groupid" value=<%=request("Groupid")%>>
		<input type="hidden" name="Groupname" value=<%=request("GroupName")%>>
		<tr> 
			<th colspan=2 height=25 class="tableHeaderText"><%=request("Groupname")%>->增加新的插件 | <a href="?action=plusinfo&Groupid=<%=request("Groupid")%>&Groupname=<%=request("Groupname")%>"><b>插件列表</b></a></th>
		</tr>
	  <tr> 
	    <td width="40%" height=25 class=forumrow align=right>插件名称&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input type="text" name="name" size=40></td>
	  </tr>
	  <tr> 
	    <td width="40%" height=25 class=forumrow align=right>插件连接文件&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input type="text" name="url" size=40></td>
	  </tr>
	  <tr> 
	    <td width="40%" height=25 class=forumrow align=right>插件简介&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input type="text" name="readme" size=40></td>
	  </tr>
	  <tr> 
	    <td width="40%" height=25 class=forumrow align=right>本插件是否提供后台管理页面&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input name=ifadmin type=radio value=0 checked>&nbsp;未提供&nbsp;&nbsp;<input type=radio name=ifadmin value=1>&nbsp;提　供</td>
	  </tr>
	  <tr> 
	    <td width="40%" height=25 class=forumrow align=right>插件后台管理连接文件&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input type="text" name="adminurl" size=40></td>
	  </tr>
		<tr> 
	    <td width="40%" height=25 class=forumrow align=right>本插件是否在新页面中打开&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input name=wind type=radio value=1 checked>&nbsp;本页面&nbsp;&nbsp;<input type=radio name=wind value=2>&nbsp;新页面 (标准)&nbsp;&nbsp;<input type="radio" name="wind" value="3">&nbsp;新页面 (限制)</td>
	  </tr>
    <tr> 
      <td width="40%" height=25 class=forumrow align=right>新窗口大小<br>(只有上面选择了限制新页面才需填写)&nbsp;</td>
      <td width="60%" class=forumrow align=left>宽度：&nbsp;<input name="width" type="text" size="10">&nbsp;&nbsp;高度：&nbsp;<input name="height" type="text" size="10"></td>
	  </tr>
	  <tr>
	    <th height="25" colspan="2" class="tableHeaderText">插件使用资格设置</th>
	  </tr>
	  <tr> 
	    <td width="40%" height=25 class=forumrow align=right>本插件是否允许被锁定和屏蔽用户访问&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input name=ifsd type=radio value=0 checked>&nbsp;禁　止&nbsp;&nbsp;<input type=radio name=ifsd value=1>&nbsp;允　许</td>
	  </tr>
		<tr> 
	    <td width="40%" height=25 class=forumrow align=right>本插件是否开启限时开关功能&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input name=iftime type=radio value=0 checked>&nbsp;关　闭&nbsp;&nbsp;<input type=radio name=iftime value=1>&nbsp;开　启</td>
	  </tr>
	  <tr> 
      <td width="40%" height=25 class=forumrow align=right>插件限制时间段(只有限时功能开启才有效)<br>目前支持2段，请使用|分隔时间&nbsp;</td>
      <td width="60%" class=forumrow align=left>&nbsp;第一时间段：&nbsp; <input name="ftime" type="text" size="10" value="8|11">&nbsp;&nbsp;第二时间段：&nbsp;<input name="stime" type="text" size="10" value="14|16"></td>
	  </tr>
	  <tr> 
	    <td width="40%" height=25 class=forumrow align=right>本插件是否开启积分限制功能&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input name=ifjf type=radio value=0 checked>&nbsp;关　闭&nbsp;&nbsp;<input type=radio name=ifjf value=1>&nbsp;开　启</td>
	  </tr>
	  <tr> 
      <td width="40%" height=25 class=forumrow align=right>积分限制范围(仅积分限制功能开启才有效)<br>使用|分隔上下限，上限不限制请输入-1&nbsp;</td>
      <td width="60%" class=forumrow align=left>&nbsp;文章：&nbsp;<input name="ifjfwz" type="text" size="8" value="0|-1">&nbsp;&nbsp;魅力：&nbsp;<input name="ifjfml" type="text" size="8" value="0|-1">&nbsp;&nbsp;经验：&nbsp;<input name="ifjfjy" type="text" size="8" value="0|-1"><br>&nbsp;威望：&nbsp;<input name="ifjfww" type="text" size="8" value="0|-1">&nbsp;&nbsp;金钱：&nbsp;<input name="ifjfjq" type="text" size="8" value="0|-1"></td>
	  </tr>
	  <tr>
	    <th height="25" colspan="2" class="tableHeaderText">插件使用设置</th>
	  </tr>
	  <tr> 
	    <td width="40%" height=25 class=forumrow align=right>本插件是否开启增减积分功能&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input name=ifjfj type=radio value=0 checked>&nbsp;关　闭&nbsp;&nbsp;<input type=radio name=ifjfj value=1>&nbsp;开　启</td>
	  </tr>
	  <tr> 
      <td width="40%" height=25 class=forumrow align=right>增减积分值(只有增减积分功能开启才有效)<br>减少积分请加上“-”号&nbsp;</td>
      <td width="60%" class=forumrow align=left>&nbsp;文章：&nbsp;<input name="ifjfjwz" type="text" size="8" value="0">&nbsp;&nbsp;魅力：&nbsp;<input name="ifjfjml" type="text" size="8" value="0">&nbsp;&nbsp;经验：&nbsp;<input name="ifjfjjy" type="text" size="8" value="0"><br>&nbsp;威望：&nbsp;<input name="ifjfjww" type="text" size="8" value="0">&nbsp;&nbsp;金钱：&nbsp;<input name="ifjfjjq" type="text" size="8" value="0"></td>
	  </tr>
	  <tr> 
	    <td width="40%" height=25 class=forumrow align=right>本插件是否启用&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input name=ifenable type=radio value=0 checked>&nbsp;启　用&nbsp;&nbsp;<input type=radio name=ifenable value=1>&nbsp;暂时停用</td>
	  </tr>
	  <tr> 
	    <td height="25" colspan="2" class=forumrow align=center><input type="submit" name="Submit" value="添 加"></td>
	  </tr>
	</table>
	</form>
<%end sub

sub savenew()
	if Request("url")<>"" and Request("readme")<>"" and request("name")<>"" then
		if request("wind")=3 then
			if request("width")="" or request("height")="" then 
				body=body+"<br>"+"请输入完整插件信息。"
				exit sub
			elseif not (isnumeric(request("width")) and isnumeric(request("height"))) then
				body=body+"<br>"+"请输入的长或宽不是有效数字。"
				exit sub
			end if
		end if
		dim linknum,windows
		set rs= server.createobject ("adodb.recordset")
		sql = "select * from plus order by plusid desc"
		rs.Open sql,connplus,1,3
		if rs.eof and rs.bof then
			linknum=1
		else
			linknum=rs("plusid")+1
		end if
		if Request.Form ("wind")=2 then
			windows="target=_blank"
		elseif request("wind")=3 then
			windows="limit"
		else
			windows=" "
		end if
		rs.AddNew 
		rs("plusid")=linknum
		rs("plusname") = Trim(Request.Form ("name"))
		rs("readme") =  Trim(Request.Form ("readme"))
		rs("url") = Request.Form ("url")
		rs("window")=windows
		if windows="limit" then
			rs("width")=request("width")
			rs("height")=request("height")
		else
			rs("width")=0
			rs("height")=0
		end if
		rs("Groupid")=request("Groupid")
		rs("Groupname")=Trim(Request.Form ("Groupname"))
		rs("config")=trim(request.form("ifenable"))+","+trim(request.form("iftime"))+","+trim(request.form("ftime"))+","+trim(request.form("stime"))+","+trim(request.form("ifadmin"))+","+trim(request.form("adminurl"))+","+trim(request.form("ifjf"))+","+trim(request.form("ifjfwz"))+","+trim(request.form("ifjfml"))+","+trim(request.form("ifjfjy"))+","+trim(request.form("ifjfww"))+","+trim(request.form("ifjfjq"))+","+trim(request.form("ifsd"))+","+trim(request.form("ifjfj"))+","+trim(request.form("ifjfjwz"))+","+trim(request.form("ifjfjml"))+","+trim(request.form("ifjfjjy"))+","+trim(request.form("ifjfjww"))+","+trim(request.form("ifjfjjq"))
		rs.Update 
		rs.Close
		set rs=nothing
		response.redirect "?action=plusinfo&Groupid="&request("Groupid")&"&Groupname="&request("Groupname")
	else
		body=body+"<br>"+"请输入完整插件信息。"
	end if
end sub

sub editlink()
	set rs= server.createobject ("adodb.recordset")
	sql = "select * from plus where plusid="&Request("plusid")
	rs.open sql,connplus,1,1
	Plus_Config=split(rs("config"),",")%>
	<form action="?action=savedit" method=post>
	<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center class="tableBorder">
		<input type="hidden" name="Groupid" value=<%=request("Groupid")%>>
		<input type="hidden" name="Groupname" value=<%=request("GroupName")%>>
		<input type="hidden" name="plusid" value=<%=Request("plusid")%>>
		<tr> 
			<th colspan=2 height=25 class="tableHeaderText"><%=request("Groupname")%>->编辑插件 | <a href="?action=plusinfo&Groupid=<%=request("Groupid")%>&Groupname=<%=request("Groupname")%>"><b>插件列表</b></a></th>
		</tr>
	  <tr> 
	    <td width="40%" height=25 class=forumrow align=right>插件名称&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input type="text" name="name" size=40 value=<%=rs("plusname")%>></td>
	  </tr>
	  <tr> 
	    <td width="40%" height=25 class=forumrow align=right>插件连接文件&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input type="text" name="url" size=40 value=<%=rs("url")%>></td>
	  </tr>
	  <tr> 
	    <td width="40%" height=25 class=forumrow align=right>插件简介&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input type="text" name="readme" size=40 value=<%=rs("readme")%>></td>
	  </tr>
	  <tr> 
	    <td width="40%" height=25 class=forumrow align=right>本插件是否提供后台管理页面&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input name=ifadmin type=radio value=0 <%if Plus_Config(4)="0" then response.write "checked"%>>&nbsp;未提供&nbsp;&nbsp;<input type=radio name=ifadmin value=1 <%if Plus_Config(4)="1" then response.write "checked"%>>&nbsp;提　供</td>
	  </tr>
	  <tr> 
	    <td width="40%" height=25 class=forumrow align=right>插件后台管理连接文件&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input type="text" name="adminurl" size=40 value="<%=Plus_Config(5)%>"></td>
	  </tr>
		<tr> 
	    <td width="40%" height=25 class=forumrow align=right>本插件是否在新页面中打开&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input name=wind type=radio value=1 <%if rs("window")=" " then response.write "checked"%>>&nbsp;本页面&nbsp;&nbsp;<input type=radio name=wind value=2 <%if rs("window")="target=_blank" then response.write "checked"%>>&nbsp;新页面 (标准)&nbsp;&nbsp;<input type="radio" name="wind" value="3" <%if rs("window")="limit" then response.write "checked"%>>&nbsp;新页面 (限制)</td>
	  </tr>
    <tr> 
      <td width="40%" height=25 class=forumrow align=right>新窗口大小<br>(只有上面选择了限制新页面才需填写)&nbsp;</td>
      <td width="60%" class=forumrow align=left>宽度：&nbsp;<input name="width" type="text" size="10" <%if rs("window")="limit" then response.write "value="&rs("width")%>>&nbsp;&nbsp;高度：&nbsp;<input name="height" type="text" size="10" <%if rs("window")="limit" then response.write "value="&rs("height")%>></td>
	  </tr>
	  <tr>
	    <th height="25" colspan="2" class="tableHeaderText">插件使用资格设置</th>
	  </tr>
	  <tr> 
	    <td width="40%" height=25 class=forumrow align=right>本插件是否允许被锁定和屏蔽用户访问&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input name=ifsd type=radio value=0 <%if Plus_Config(12)="0" then response.write "checked"%>>&nbsp;禁　止&nbsp;&nbsp;<input type=radio name=ifsd value=1 <%if Plus_Config(12)="1" then response.write "checked"%>>&nbsp;允　许</td>
	  </tr>
		<tr> 
	    <td width="40%" height=25 class=forumrow align=right>本插件是否开启限时开关功能&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input name=iftime type=radio value=0 <%if Plus_Config(1)="0" then response.write "checked"%>>&nbsp;关　闭&nbsp;&nbsp;<input type=radio name=iftime value=1 <%if Plus_Config(1)="1" then response.write "checked"%>>&nbsp;开　启</td>
	  </tr>
	  <tr> 
      <td width="40%" height=25 class=forumrow align=right>插件限制时间段(只有限时功能开启才有效)<br>目前支持2段，请使用|分隔时间&nbsp;</td>
      <td width="60%" class=forumrow align=left>&nbsp;第一时间段：&nbsp; <input name="ftime" type="text" size="10" value="<%=Plus_Config(2)%>">&nbsp;&nbsp;第二时间段：&nbsp;<input name="stime" type="text" size="10" value="<%=Plus_Config(3)%>"></td>
	  </tr>
	  <tr> 
	    <td width="40%" height=25 class=forumrow align=right>本插件是否开启积分限制功能&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input name=ifjf type=radio value=0 <%if Plus_Config(6)="0" then response.write "checked"%>>&nbsp;关　闭&nbsp;&nbsp;<input type=radio name=ifjf value=1 <%if Plus_Config(6)="1" then response.write "checked"%>>&nbsp;开　启</td>
	  </tr>
	  <tr> 
      <td width="40%" height=25 class=forumrow align=right>积分限制范围(仅积分限制功能开启才有效)<br>使用|分隔上下限，上限不限制请输入-1&nbsp;</td>
      <td width="60%" class=forumrow align=left>&nbsp;文章：&nbsp;<input name="ifjfwz" type="text" size="8" value="<%=Plus_Config(7)%>">&nbsp;&nbsp;魅力：&nbsp;<input name="ifjfml" type="text" size="8" value="<%=Plus_Config(8)%>">&nbsp;&nbsp;经验：&nbsp;<input name="ifjfjy" type="text" size="8" value="<%=Plus_Config(9)%>"><br>&nbsp;威望：&nbsp;<input name="ifjfww" type="text" size="8" value="<%=Plus_Config(10)%>">&nbsp;&nbsp;金钱：&nbsp;<input name="ifjfjq" type="text" size="8" value="<%=Plus_Config(11)%>"></td>
	  </tr>
	  <tr>
	    <th height="25" colspan="2" class="tableHeaderText">插件使用设置</th>
	  </tr>
	  <tr> 
	    <td width="40%" height=25 class=forumrow align=right>本插件是否开启增减积分功能&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input name=ifjfj type=radio value=0 <%if Plus_Config(13)="0" then response.write "checked"%>>&nbsp;关　闭&nbsp;&nbsp;<input type=radio name=ifjfj value=1 <%if Plus_Config(13)="1" then response.write "checked"%>>&nbsp;开　启</td>
	  </tr>
	  <tr> 
      <td width="40%" height=25 class=forumrow align=right>增减积分值(只有增减积分功能开启才有效)<br>减少积分请加上“-”号&nbsp;</td>
      <td width="60%" class=forumrow align=left>&nbsp;文章：&nbsp;<input name="ifjfjwz" type="text" size="8" value="<%=Plus_Config(14)%>">&nbsp;&nbsp;魅力：&nbsp;<input name="ifjfjml" type="text" size="8" value="<%=Plus_Config(15)%>">&nbsp;&nbsp;经验：&nbsp;<input name="ifjfjjy" type="text" size="8" value="<%=Plus_Config(16)%>"><br>&nbsp;威望：&nbsp;<input name="ifjfjww" type="text" size="8" value="<%=Plus_Config(17)%>">&nbsp;&nbsp;金钱：&nbsp;<input name="ifjfjjq" type="text" size="8" value="<%=Plus_Config(18)%>"></td>
	  </tr>
	  <tr> 
	    <td width="40%" height=25 class=forumrow align=right>本插件是否启用&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input name=ifenable type=radio value=0 <%if Plus_Config(0)="0" then response.write "checked"%>>&nbsp;启　用&nbsp;&nbsp;<input type=radio name=ifenable value=1 <%if Plus_Config(0)="1" then response.write "checked"%>>&nbsp;暂时停用</td>
	  </tr>
	  <tr> 
	    <td height="25" colspan="2" class=forumrow align=center><input type="submit" name="Submit" value="修 改"></td>
	  </tr>
	</table>
	</form>
	<%rs.close
	set rs=nothing
end sub

sub savedit()
	if Request("url")<>"" and Request("readme")<>"" and request("name")<>"" then
		if request("wind")=3 then
			if request("width")="" or request("height")="" then 
				body=body+"<br>"+"请输入完整插件信息。"
				exit sub
			elseif not (isnumeric(request("width")) and isnumeric(request("height"))) then
				body=body+"<br>"+"请输入的长或宽不是有效数字。"
				exit sub
			end if
		end if
		set rs= server.createobject ("adodb.recordset")
		sql = "select * from plus where plusid="&request("plusid")
		rs.Open sql,connplus,1,3
		if rs.eof and rs.bof then
			body=body+"<br>"+"错误，没有找到插件。"
		else
			if Request.Form("wind")=2 then
				rs("window")="target=_blank"
			elseif request("wind")=3 then
				rs("window")="limit"
				rs("width")=request("width")
				rs("height")=request("height")
			else
				rs("window")=" "
			end if	
			rs("plusname") = Trim(Request.Form ("name"))
			rs("readme") =  Trim(Request.Form ("readme"))
			rs("url") = Request.Form ("url")
			rs("config")=trim(request.form("ifenable"))+","+trim(request.form("iftime"))+","+trim(request.form("ftime"))+","+trim(request.form("stime"))+","+trim(request.form("ifadmin"))+","+trim(request.form("adminurl"))+","+trim(request.form("ifjf"))+","+trim(request.form("ifjfwz"))+","+trim(request.form("ifjfml"))+","+trim(request.form("ifjfjy"))+","+trim(request.form("ifjfww"))+","+trim(request.form("ifjfjq"))+","+trim(request.form("ifsd"))+","+trim(request.form("ifjfj"))+","+trim(request.form("ifjfjwz"))+","+trim(request.form("ifjfjml"))+","+trim(request.form("ifjfjjy"))+","+trim(request.form("ifjfjww"))+","+trim(request.form("ifjfjjq"))
			rs.Update
		end if 
		rs.Close
		set rs=nothing
		response.redirect "?action=plusinfo&Groupid="&request("Groupid")&"&Groupname="&request("Groupname")
	else
		body=body+"<br>"+"请输入完整插件信息。"
	end if
end sub

sub del
	dim plusid
	plusid = request("plusid")
	sql="delete from plus where plusid="+plusid
	connplus.Execute(sql)
	response.redirect "?action=plusinfo&Groupid="&request("Groupid")&"&Groupname="&request("Groupname")
end sub

sub OrdersUp()
	dim oldid,newid
	set rs= server.createobject ("adodb.recordset")
	sql="select top 2 plusid from plus where plusid<="&request("plusid")&" and groupid="&request("groupid")&" order by plusid desc"
	rs.open sql,connplus,1,3
	if not (rs.eof or rs.bof) then
		oldid=rs("plusid")
		rs.movenext
		if not (rs.eof or rs.bof) then
			newid=rs("plusid")
			rs("plusid")=oldid
			rs.update
			rs.MovePrevious
			rs("plusid")=newid
			rs.update
		end if
	end if
	rs.close
	set rs=nothing
	response.redirect "?action=plusinfo&Groupid="&request("Groupid")&"&Groupname="&request("Groupname")
end sub

sub OrdersDown()
	dim oldid,newid
	set rs= server.createobject ("adodb.recordset")
	sql="select top 2 plusid from plus where plusid>="&request("plusid")&" and groupid="&request("groupid")&" order by plusid"
	rs.open sql,connplus,1,3
	if not (rs.eof or rs.bof) then
		oldid=rs("plusid")
		rs.movenext
		if not (rs.eof or rs.bof) then
			newid=rs("plusid")
			rs("plusid")=oldid
			rs.update
			rs.MovePrevious
			rs("plusid")=newid
			rs.update
		end if
	end if
	rs.close
	set rs=nothing
	response.redirect "?action=plusinfo&Groupid="&request("Groupid")&"&Groupname="&request("Groupname")
end sub

sub tj()
	set rs= server.createobject ("adodb.recordset")
	sql = " select * from plus order by getnum desc"
	rs.open sql,connplus,1,1%> 
	<table width="100%" border="0" cellspacing="1" cellpadding="4"  align=center class="tableBorder">
	  <tr> 
	    <th width="87" height=25  class="tableHeaderText">插件ID</td>
	    <th width="144" class="tableHeaderText">插件名称</td>
	    <th width="144" class="tableHeaderText">是否启用</td>
	    <th width="*" class="tableHeaderText">总访问量</td>
	  </tr>
	  <%if not rs.eof then
			do while not rs.eof%>
				<tr align=center>
					<td height=25 class=forumrow width="87"><font color=<%=forum_body(8)%>><%=rs("plusid")%></font></td>
					<td  class=forumrow width="144"><%=rs("plusname")%></td>
	    		<%Plus_Config=split(rs("config"),",")
	      	if Plus_Config(0)="1" then
						Plus_Config(2)="禁用"
					else
						Plus_Config(2)="启用"
					end if%>
	    		<td class=forumrow width="144"><%=Plus_Config(2)%></td>
					<td class=forumrow width="*"><b><%=rs("getnum")%></b></td>
				</tr>
				<%rs.movenext
			loop
		else
			response.write "<tr><td colspan='4' class=forumrow align=center>目前还没有插件！</td></tr>"
		end if
		rs.Close
		set rs=nothing%>
	</table>
<%end sub%>























