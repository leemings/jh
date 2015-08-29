<table cellspacing=1 width="100%" align=center bgcolor=<%=Forum_Body(27)%>>
	<tr>
		<th valign=middle height=25 width="100%">热门推荐</th>
	</tr>
	<tr>
		<td class=tablebody1 style="word-break:break-all"><br><%
			dim h
			sql="select top "&hotnum&" id,showname,hits from download where hots=1"
			sql=sql&" order by hits desc"
			rss.open sql,conndown,1,1
			if rss.eof and rss.bof then
				%>&nbsp;没有推荐<br><%
			else
				h=0
				do while not rss.eof and h<hotnum
					%>&nbsp;<font color=red>◆</font> <a href="z_down_list.asp?id=<%=rss("id")%>"><%=rss("showname")%></a><font color=red>[<%=rss("hits")%>]</font><br><%
					rss.movenext
					h=h+1
				loop
			end if
			rss.close
		%><br></td>
	</tr>
</table>