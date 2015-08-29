<table cellspacing=1 width="100%" align=center bgcolor=<%=Forum_Body(27)%>>
	<tr>
		<th valign=middle height=25 width="100%">今日下载Top<%=daytop%></th>
	</tr>
	<tr>
		<td class=tablebody1 style="word-break:break-all"><br><%
			dim tdate,d
			tdate=year(Now()) & "-" & month(Now()) & "-" & day(Now())
			set rss=server.createobject("adodb.recordset")
			sql="select top "&daytop&" id,showname,dayhits from download where lasthits=date() and stop=0 and dayhits>0 order by dayhits desc"
			rss.open sql,conndown,1,1
			if rss.eof and rss.bof then
				%>&nbsp;本日没有下载<br><%
			else
				d=0
				do while (not rss.eof and d<daytop)
					%>&nbsp;<font color=red>☆</font> <a href="z_down_list.asp?id=<%=rss("id")%>"><%=rss("showname")%></a><font color=red>[<%=rss("dayhits")%>]</font><br><%
					rss.movenext
					d=d+1
				loop
			end if
		%><br></td>
	</tr>
</table>