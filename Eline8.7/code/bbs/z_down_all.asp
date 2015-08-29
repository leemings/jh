<table cellspacing=1 width="100%" align=center bgcolor=<%=Forum_Body(27)%>>
	<tr>
		<th valign=middle height=25 width="100%">累计下载Top<%=alltop%></th>
	</tr>
	<tr>
		<td class=tablebody1 style="word-break:break-all"><br><%
			dim im
			sql="select top "&alltop&" * from download order by hits desc"
			set rss=server.createobject("adodb.recordset")
			rss.open sql,conndown,1,1
			if rss.eof and rss.bof then
				%>&nbsp;没有下载<br><%
			else
				im=0
				do while not rss.eof and im<alltop
					%>&nbsp;<font color=red>☆</font> <a href="z_down_list.asp?id=<%=rss("id")%>"><%=rss("showname")%></a><font color=red>[<%=rss("hits")%>]</font><br><%
					rss.movenext
					im=im+1
				loop
			end if
			rss.close
		%><br></td>
	</tr>
</table>
