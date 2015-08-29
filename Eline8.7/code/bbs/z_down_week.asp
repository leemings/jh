<table cellspacing=1 width="100%" align=center bgcolor=<%=Forum_Body(27)%>>
	<tr>
		<th valign=middle height=25 width="100%">本周下载Top<%=weektop%></th>
	</tr>
	<tr>
		<td class=tablebody1 style="word-break:break-all"><br><%
			dim OldWeek,NewWeek,u
			OldWeek = WeekDay(Date())-1
			If OldWeek = 0 Then OldWeek = 7
			OldWeek = Date()-OldWeek
			NewWeek = Date()+(9-WeekDay(Date()))
			set rss=server.createobject("adodb.recordset")
			sql="select top "&weektop&" id,showname,weekhits from download where "
			sql=sql&"  (lasthits < #" & NewWeek & "#) And (lasthits > #" & OldWeek & "#) and stop=0 and weekhits>0 "
			sql=sql&" order by weekhits desc"
			rss.open sql,conndown,0,1
			if rss.eof and rss.bof then
				%>&nbsp;本周没有下载<br><%
			else
				u=0
				do while not rss.eof and u<weektop
					%>&nbsp;<font color=red>☆</font> <a href="z_down_list.asp?id=<%=rss("id")%>"><%=rss("showname")%></a><font color=red>[<%=rss("weekhits")%>]</font><br><%
					rss.movenext
					u=u+1
				loop
			end if
			rss.close
		%><br></td>
	</tr>
</table>