<!--#include file="conn.asp"-->
<!--#include file="z_photo_conn.asp"-->
<!-- #include file="inc/const.asp" -->
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Forum_css.asp"-->
<script language="JavaScript" type="text/JavaScript">
function picsize()
{
var a,b
a=document.all["z_photo_pic"].width;
b=document.all["z_photo_pic"].height;
if (a>b)
{
if (document.all["z_photo_pic"].width>500)
document.all["z_photo_pic"].width=500;
}
else {
if (document.all["z_photo_pic"].height>400)
document.all["z_photo_pic"].height=400;
}
}
function showpe()
{
document.all["p_edit"].style.display=""
}
function closepe()
{
document.all["p_edit"].style.display="none"
}
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" <%=Forum_body(11)%>>
<%if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>您没有进入个人相册的权限，请先登录或者同管理员联系。"
	call dvbbs_error()
else
	dim tmp_id
	dim pho_url,pho_title,pho_brief,pho_adtime,username,pho_state,pho_shared,pho_group,pho_count
	dim pho_size
	tmp_id=clng(request.QueryString("pho_id"))
	dim menu
	menu=request.QueryString("menu")
	select case menu
	case 1
		call del()				'删除照片
	case 2
		call share()			'设置照片共享属性
	case 3
		call movepho()			'移动照片
	case 4
		call newmess()			'添加新的留言
	case 5
		call messdel()			'删除留言
	case 6
		call upinfo()			'更新照片信息
	case else 
	end select
	set rs=server.createobject("adodb.recordset")
	sql="select * from z_photo where pho_id = "&tmp_id
	rs.open sql,z_photo_cn,1,3
	if rs.eof then
		Response.Write "没有这张照片！"
	else
		rs("pho_count")=rs("pho_count")+1
		rs.update
		username=rs("username")
		pho_url=rs("pho_url")
		pho_adtime=rs("pho_adtime")
		pho_adtime=formatdatetime(pho_adtime,2)
		pho_title=rs("pho_title")
		pho_brief=rs("pho_brief")
		pho_group=rs("pho_group")
		pho_state=rs("state")
		pho_shared=rs("shared")
		pho_count=rs("pho_count")
		pho_size=rs("pho_size")
		rs.close%>
		<form name="form1" method="post" action="z_photo_view1.asp?pho_id=<%=tmp_id%>&pho_group=<%=pho_group%>&menu=3">
		<table width="97%" align="center" cellpadding="3" cellspacing="1" class=tableborder1>
			<tr> 
				<th colspan="3" height=25>-=> 标题：<%= pho_title %></th>
			</tr>
			<tr> 
				<td height=21 colspan="3" valign=top class=tablebody2 align="center"><a href="<%= pho_url %>"><img id="z_photo_pic" src="<%= pho_url %>" alt="点击查看原始图片" onload="picsize()" border="0"></a></td>
			</tr>
			<tr> 
				<td height="20" colspan="3" valign=top class=tablebody2 align="center"><% 
					dim u_id,d_id
					set rs=server.createobject("adodb.recordset")
					sql="select pho_id from z_photo where state = 'v' and pho_id > "&tmp_id&" and pho_group ='"&pho_group&"' order by pho_id"
					rs.open sql,z_photo_cn,1,3
					if not rs.eof then
						u_id=rs("pho_id")
						%><a href="z_photo_view1.asp?pho_id=<%=u_id%>"> 上一张  </a><%
					end if 
					rs.close
					set rs=server.createobject("adodb.recordset")
					sql="select pho_id from z_photo where state = 'v' and pho_id < "&tmp_id&" and pho_group ='"&pho_group&"' order by pho_id desc"
					rs.open sql,z_photo_cn,1,3
					if not rs.eof then
						d_id=rs("pho_id")
						%><a href="z_photo_view1.asp?pho_id=<%=d_id%>">  下一张 </a><%
					end if 
					rs.close
				%></td>
				</tr>
				<%if username=membername or master or not session("flag")="" then %>
					<tr> 
						<td height=21 colspan="2" valign=top class=tablebody2> 转移至 <select name="pho_groups" id="pho_groups"><% 
						dim pho_group_id,per_name
						set rs=server.createobject("adodb.recordset")
						sql="select * from z_photo_group where (pho_state = 'v' and pho_all='y') "
						rs.open sql,z_photo_cn,1,3
						do while not rs.eof 
							pho_group_id=rs("pho_group_id")
							per_name=rs("pho_group_name")
							%><option selected value="<%= pho_group_id %>" ><%= per_name %></option><%
							rs.movenext 
						loop
						rs.close
						set rs=server.createobject("adodb.recordset")
						sql="select * from z_photo_group where (pho_state = 'v' and username='"&membername&"' and pho_all='n') "
						rs.open sql,z_photo_cn,1,3
						if not rs.eof then
							do while not rs.eof 
								pho_group_id=rs("pho_group_id")
								per_name=rs("pho_group_name")
								%><option value="<%= pho_group_id %>" ><%= per_name %></option><%
								rs.movenext 
							loop
							rs.close
							%></select><%
						else 
							Response.Write "</select>"
							Response.Write "你还没有建立个人相册，请到相册组管理页面中添加！"
						end if
					%><input name="go" type="submit" id="go2" value="转移"><input name="pho_size" type="hidden" id="pho_size" value="<%= pho_size %>"></td>
					<td width="50%" valign=top class=tablebody2 align="center"><input type="button" name="Button" value="编辑信息"  onClick="showpe()"><input name="del" type="button" id="del" value="删除照片" onClick="window.location='z_photo_view1.asp?pho_id=<%=tmp_id%>&pho_group=<%=pho_group%>&pho_size=<%=pho_size%>&menu=1'"></td>
				</tr>
			<% end if %>
			<tr> 
				<td width="13%" valign=top class=tablebody2> 引用地址：</td>
				<td height=20 colspan="2" valign=top class=tablebody1><%= pho_url %></td>
			</tr>
			<tr> 
				<td height="22" class=tablebody2>添加时间：</td>
				<td width="37%" class=tablebody1><%= pho_adtime %> by <%= username %></td>
				<td class=tablebody1 align="center"> 点击数：<font color="red"><%= pho_count %></font></td>
			</tr>
			<tr> 
				<td width="13%" valign=top class=tablebody2> 照片说明：</td>
				<td height=20 colspan="2" valign=top class=tablebody1><%= pho_brief %></td>
			</tr>    
		</table>
		</form>
		<form name="form3" method="post" action="z_photo_view1.asp?pho_id=<%=tmp_id%>&pho_group=<%=pho_group%>&menu=6">
		<table id="p_edit" width="97%" border="0" align="center" cellpadding="3" cellspacing="0" class="tableBorder1" style="display:none">
			<tr> 
				<th height="25" colspan="3" align="left">-=>编辑照片信息</th>
			</tr>
			<tr> 
				<td width="13%" class="TableBody2">照片标题：</td>
				<td colspan="2" class="TableBody2"> <input name="p_title" type="text" id="p_title" value="<%= pho_title %>" size="50"></td>
			</tr>
			<tr> 
				<td class="TableBody2">照片说明：</td>
				<td colspan="2" class="TableBody2"> <textarea name="p_brief" cols="50" wrap="VIRTUAL" id="p_brief"><%= pho_brief %></textarea> <input type="submit" name="Submit" value="更新"> <input type="button" name="Submit2" value="算了" onClick="closepe()"></td>
			</tr>
		</table>
		</form>
		<table width="97%" border="0" align="center">
			<tr> 
				<th height="25" align="left">-=>大家来说两句！</th>
			</tr>
			<form name="form2" method="post" action="z_photo_view1.asp?pho_id=<%=tmp_id%>&menu=4">
				<tr>
					<td class=tablebody2><textarea name="message" cols="70" rows="6" wrap="VIRTUAL" id="textarea"></textarea><input name="say" type="submit" id="say2" value="说两句！" ><input name="bye" type="button" id="bye2" value="关闭" onClick="window.close()"></td>
				</tr>
			</form>
			<tr> 
  			<td><% 
					dim pho_mess,pho_m_user,pho_m_time,mess_id
					set rs=server.createobject("adodb.recordset")
					sql="select * from z_photo_mess where (state = 'v' and pho_id ="&tmp_id&") order by addtime desc"
					rs.open sql,z_photo_cn,1,3
					if rs.eof then
						Response.Write "目前还没有留言！"
					else
						do while not rs.eof 
							mess_id=rs("id")
							pho_m_user=rs("username")
							pho_mess=rs("message")
							pho_m_time=rs("addtime")
							rs.movenext%>
							<table width="80%" border="0" align="center" cellpadding="5" cellspacing="1" class=tableborder1>
								<tr> 
									<td class=tablebody1><p><%=pho_mess%></p><p>作者：<%= pho_m_user %> 时间：<%= pho_m_time %> <%
										if username=membername then
											%><a href="z_photo_view1.asp?pho_id=<%=tmp_id%>&mess_id=<%=mess_id%>&menu=5"><font color="red">删除</font></a><%
										end if
									%></p></td>
								</tr>
							</table>
						<%loop
					end if
					rs.close
				%></td>
			</tr>
		</table>
	<%end if 
end if %>
</body>
</html>
<% 
call close_z_photo_cn()
conn.close
set conn=nothing

'----主程序段落------
'----right:clasky.com,seasky7----------------
'--------------------------删除照片---------------------------------
sub del()
	pho_group=request.QueryString("pho_group")
	pho_size=request.QueryString("pho_size")
	set rs=server.createobject("adodb.recordset")
	sql="select * from z_photo where pho_id = "&tmp_id
	rs.open sql,z_photo_cn,1,3
	rs("state")="d"
	rs.update
	rs.close	
	call down_group(pho_group,pho_size)
	Response.Write "<script>"
	Response.Write "alert('已经成功删除照片！"&pho_size&"');"
	response.write "opener.location.reload();"
	Response.Write "window.close();</script>"
end sub
'------------------------设置共享---------------------------
sub share()
	dim show
	set rs=server.createobject("adodb.recordset")
	sql="select * from z_photo where pho_id = "&tmp_id
	rs.open sql,z_photo_cn,1,3
	if rs("shared")="n" then
		rs("shared")="y"
		show="1"
	else
		rs("shared")="n"
		show="0"
	end if
	rs.update
	rs.close	
	Response.Write "<script>"
	Response.Write "alert('已经成功设置照片！');"
	response.write "opener.top.location.reload();"
	Response.Write "</script>"
end sub
'------------------移动程序----------------
sub movepho()
	dim tpho_groups,pho_groups,tshared,pho_size
	tpho_groups=request.form("pho_groups")
	pho_group=request.QueryString("pho_group")
	pho_size=request.form("pho_size")
	set rs=server.createobject("adodb.recordset")
	sql="select * from z_photo where pho_id = "&tmp_id
	rs.open sql,z_photo_cn,1,3
	rs("pho_group")=tpho_groups
	rs.update
	rs.close	
	call up_group(tpho_groups,pho_size)
	call down_group(pho_group,pho_size)
	Response.Write "<script>"
	Response.Write "alert('已经成功转移照片！"&pho_size&"');"
	response.write "opener.top.location.reload();"
	Response.Write "</script>"
end sub
'------------新的留言---------------
sub newmess()
	dim message
	message=request.form("message")
	if not message="" then
		z_photo_cn.execute("insert into z_photo_mess (pho_id,username,message) values ('"&tmp_id&"','"&membername&"','"&message&"')")
		Response.Write "<script>"
		Response.Write "alert('您已成功发言！');"
		Response.Write "</script>"
	else
	end if
end sub
'--------------------------删除留言---------------------------------
sub messdel()
	mess_id=request.QueryString("mess_id")
	set rs=server.createobject("adodb.recordset")
	sql="select * from z_photo_mess where id = "&mess_id
	rs.open sql,z_photo_cn,1,3
	rs("state")="d"
	rs.update
	rs.close	
	Response.Write "<script>"
	Response.Write "alert('已经成功删除留言！');"
	Response.Write "</script>"
end sub
'-----------------更新相册统计信息-增加-------------------
sub up_group(d,s)
	d=clng(d)
	s=clng(s)
	set rs=server.createobject("adodb.recordset")
	sql="select pho_group_totle,pho_group_size from z_photo_group where pho_group_id = "&d
	rs.open sql,z_photo_cn,1,3
	rs("pho_group_totle")=rs("pho_group_totle")+1
	rs("pho_group_size")=rs("pho_group_size")+s
	rs.update
	rs.close
end sub
'-----------------更新相册统计信息-减少-------------------
sub down_group(d,s)
	d=clng(d)
	s=clng(s)
	set rs=server.createobject("adodb.recordset")
	sql="select pho_group_totle,pho_group_size from z_photo_group where pho_group_id = "&d
	rs.open sql,z_photo_cn,1,3
	rs("pho_group_totle")=rs("pho_group_totle")-1
	rs("pho_group_size")=rs("pho_group_size")-s
	rs.update
	rs.close
end sub
'-----------------更新相册信息-------------------
sub upinfo()
	dim p_title,p_brief
	p_title=request.Form("p_title")
	p_brief=request.Form("p_brief")
	set rs=server.createobject("adodb.recordset")
	sql="select * from z_photo where pho_id = "&tmp_id
	rs.open sql,z_photo_cn,1,3
	rs("pho_title")=p_title
	rs("pho_brief")=p_brief
	rs.update
	rs.close
	Response.Write "<script>"
	Response.Write "alert('已经成功更新该照片信息！');"
	Response.Write "</script>"
end sub
%>
