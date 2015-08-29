<!--#include file="conn.asp"-->
<!--#include file="z_dgconn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="z_plus_check.asp"-->
<%
	dim page_count,Pcount,totalrec,mytotalrec,totalPages,currentPage
	dim abgcolor,username

	stats="所有点歌列表"
	
	call nav()
	call head_var(0,0,"点歌控制面板","z_dglistall.asp")

	if not founduser then
		errmsg=errmsg+"<br>"+"<li>您没有查阅点歌列表的权限，请先登录或与系统管理员联系"
		founderr=true
	end if
	
	if founderr then 
		call dvbbs_error()
	else	
		totalrec=connDG.execute("select count(id) from media")(0)
		mytotalrec=connDG.execute("select count(id) from media where incept='"&membername&"' or incept='全体会员'")(0)
	%>	
		<table cellpadding=3 cellspacing=1 border=0 align=center class=tableborder1>
			<tr>
				<th><a href=z_dglistall.asp><font color=white><b>所有点歌列表</b></font></a></th>
				<th><a href=z_dglistme.asp><font color=white><b>我的点歌列表</b></font></a></th>
				<th><a href=z_dgwrite.asp><font color=white><b>我要点歌</b></font></a></th>
			</tr>
			<tr>
				<td class=tablebody2 align=center valign=middle colspan="3"><a href=z_dglistall.asp><b>论坛总点歌列表</b></a>清单共有[<font color=red><b><%=totalrec%></b></font>]个,其中<a href=z_dglistme.asp><font color=green>[<b><%=membername%></b>]</font></a>的祝福清单共有[<font color=red><b><%=mytotalrec%></b></font>]个。<font color=green><直接点击歌名欣赏></font></td>
			</tr>
		</table>	
		<br>
		<table cellpadding=3 cellspacing=1 border=0 align=center class=tableborder1>
				<tr>
					<th width=80>点歌人</th><th width=80>对方姓名</th><th width=220>歌名</th><th width=120>时间</th><th width=*>祝福语</th><th width=50>操作</th> 
				</tr>	
			
	<%	
		set rs=server.createobject("adodb.recordset")
		sql="select * from media Order By sendtime Desc"
		rs.open sql,connDG,3,3				
		if rs.eof and rs.bof then
			currentpage=0
	%>	
				
			<tr><td class=tablebody2 valign=middle colspan=6>当前没有点歌列表</td></tr>
	<%		
		else

			currentPage=request.querystring("page")
			if currentpage="" or not isInteger(currentpage) then
				currentpage=1
			else
				currentpage=clng(currentpage)
				if err then
					currentpage=1
					err.clear
				end if
			end if

		rs.PageSize = 10
		rs.AbsolutePage=currentpage
		page_count=0
      	totalrec=rs.recordcount
		while (not rs.eof) and (not page_count = rs.PageSize)
			if trim(rs("incept"))=membername or trim(rs("incept"))="全体会员" then
%>
				<tr>
					<td class=tablebody2 align=center valign=middle><a href=javascript:openScript('dispuser.asp?name=<%=rs("sender")%>',650,500)><font color=blue><%=rs("sender")%></font></a></td>
					<td class=tablebody2 align=center valign=middle><%if trim(rs("incept"))=membername then%><a href=javascript:openScript('dispuser.asp?name=<%=rs("incept")%>',650,500)><font color=#CC66FF><%=rs("incept")%></font></a><%else%><font color=olive><%=rs("incept")%></font><%end if%></td>
					<td class=tablebody2 align=center valign=middle><a href='z_dgplay.asp?url=<%=replace(rs("url"),chr(32),"%20",1)%>&medianame=<%=replace(rs("medianame"),chr(32),"%20")%>' target=_blank><%=rs("medianame")%></a></td>
					<td class=tablebody2 align=center valign=middle><%=rs("sendtime")%></td>
					<td class=tablebody2 align=left valign=middle><font color=navy><%=DvBCode(rs("content"),1,2)%></font></td>
					<td class=tablebody2 align=center valign=middle><a href=z_dgdel.asp?id=<%=rs("id")%> onclick="javascript:{if(confirm('您确定执行删除操作吗?')){return true;}return false;}">[删除]</a></td>
				</tr>		
<%		
			else
%>
				<tr>
					<td class=tablebody1 align=center valign=middle><a href=javascript:openScript('dispuser.asp?name=<%=rs("sender")%>',650,500)><font color=blue><%=rs("sender")%></font></a></td>
					<td class=tablebody1 align=center valign=middle><a href=javascript:openScript('dispuser.asp?name=<%=rs("incept")%>',650,500)><font color=green><%=rs("incept")%></font></a></td>
					<td class=tablebody1 align=center valign=middle><a href='z_dgplay.asp?url=<%=replace(rs("url"),chr(32),"%20",1)%>&medianame=<%=replace(rs("medianame"),chr(32),"%20")%>' target=_blank><%=rs("medianame")%></a></td>
					<td class=tablebody1 align=center valign=middle><%=rs("sendtime")%></td>
					<td class=tablebody1 align=left valign=middle><font color=gray>******已经隐藏******</font></td>
					<td class=tablebody1 align=center valign=middle><a href=z_dgdel.asp?id=<%=rs("id")%> onclick="javascript:{if(confirm('您确定执行删除操作吗?')){return true;}return false;}">[删除]</a></td>
				</tr>
<%
			end if
	  	page_count = page_count + 1
		rs.movenext
		wend
	end if
%>
</table>
<%
	dim endpage
	Pcount=rs.PageCount
	response.write "<table border=0 cellpadding=0 cellspacing=3 width="&Forum_body(12)&" align=center>"&_
			"<tr><td valign=middle nowrap>"&_
			"<font color='#000000'>页次：<b>"&currentpage&"</b>/<b>"&Pcount&"</b>页"&_
			" 每页<b>10</b>条 共有<b>"&totalrec&"</b>条点歌</td>"&_
			"<td valign=middle nowrap><div align=right><p>分页： "
	call DispPageNum(currentpage,PCount,"""?page=","""")
		response.write "</p></div></font></td></tr></table>"
	rs.close
	set rs=nothing
'=========================
%>	
		<table cellpadding=3 cellspacing=1 border=0 align=center class=tableborder1>
			<tr>
				<th><font color=white>--== 论坛点歌台-所有点歌列表 ==--</font></th>
			</tr>
		</table>		
<%
	end if
	call CloseDB()
	call activeonline()
	call footer()
%>