<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_down_conn.asp"-->
<!--#include file="z_down_function.asp" -->
<!--#include file="z_plus_check.asp"-->
<%stats="下载中心"
call nav()
call head_var(2,0,"","")%>
<!-- #include file="z_down_forum_js.asp" -->
<%call down_nav()%>
<table cellspacing=0 cellpadding=0 align=center border="0" width="<%=Forum_body(12)%>">
	<tr>
		<td width="24%" valign="top" align="left" rowspan="2"><%
			if showset(2)=1 then
				%><!--#include file="z_down_day.asp"--><br><%
			end if
			if showset(3)=1 then
				%><!--#include file="z_down_week.asp"--><br><%
			end if
			if showset(4)=1 then
				%><!--#include file="z_down_all.asp"--><br><%
			end if
		%></td>
		<td width="52%" valign="top" align="left">
			<table cellspacing=1 width="95%" align=center bgcolor=<%=Forum_Body(27)%>>
				<tr>
					<th valign=middle height=25 width="100%">本站新进软件</th>
				</tr>
				<tr>
					<td class=tablebody1>
						<table align=center border="0" width="100%"><%
							sql="select top "&newnum&" "
							sql=sql&"download.id,download.showname,download.dateandtime,download.hits,download.classid,download.Nclassid,DNclass.Nclass "
							sql=sql&" from download,DNclass where download.stop=0 and download.Nclassid=DNclass.Nclassid "
							sql=sql&" order by download.id desc"
							rss.open sql,conndown,1,1
							if rss.eof and rss.bof then
								%>没有任何软件<%
							else
								do while not rss.eof%>
									<tr>
										<td height=20 style="word-break:break-all">&nbsp;&nbsp;[<a href='z_down_index.asp?classid=<%=rss("classid")%>&Nclassid=<%=rss("Nclassid")%>'><font color=blue><%=rss("Nclass")%></font></a>] <a href="z_down_list.asp?id=<%=rss("id")%>"><%=rss("showname")%></a>(<%if rss("dateandtime")=date() then%><font color=red><%=month(rss("dateandtime"))%>月<%=day(rss("dateandtime"))%>日</font><%else%><font color='#999999'><%=month(rss("dateandtime"))%>月<%=day(rss("dateandtime"))%>日</font><%end if%>,<font color=green><%=rss("hits")%></font>)<br></td>
									</tr>
									<%rss.movenext
								loop
							end if
							rss.close%>
	      		</table>
	      	</td>
	      </tr>
			</table>
   		<br>
		</td>
    <td width="24%" valign="top" align="left"><%
    	if showset(0)=1 then
    		%><!--#include file="z_down_allhot.asp"--><br><%
    	end if
    	if showset(1)=1 then
  			%><!--#include file="z_down_hot.asp"--><br><%
  		end if
  	%></td>
	</tr>
	<tr>
		<td valign="top" align="left" colspan="2">
			<table cellspacing=1 width="98%" align=right valign="top" bgcolor=<%=Forum_Body(27)%>>
				<tr>
					<th valign=middle height=25 width="100%">软件详细分类</th>
				</tr>
				<tr>
					<td class=tablebody1 valign="top" align="left"><br>
						<table border="0" width="100%" cellspacing="1">
							<tr>
								<td width="2%"></td>
								<td width="98%"><%
									sql="select class,classid from Dclass"
									rss.open sql,conndown,1,1
									if rss.bof and rss.eof then
										response.write "没有任何分类"
									else
										dim rs1
										do while not rss.eof
											%><IMG SRC="images/img_down/form.gif" BORDER=0>  <a href="z_down_index.asp?classid=<%=rss("classid")%>"><font color=blue><%=rss("class")%></font></a><br><%
											set rs1=server.createobject("adodb.recordset")
											sql="select Nclass,Nclassid,classid from DNclass where classid="&rss("classid")
											rs1.open sql,conndown,1,3
											do while not rs1.eof
												%><a href="z_down_index.asp?classid=<%=rs1("classid")%>&Nclassid=<%=rs1("Nclassid")%>"><%=rs1("Nclass")%></a>&nbsp;&nbsp;<%
												rs1.movenext
											loop
											rs1.close
											%><br><br><%
											rss.movenext
										loop
									end if
									rss.close
								%></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<br>
<table cellspacing=1 cellpadding=3 align=center border="0" class=tableborder1>
	<tr>
		<th height=25>搜索工具</th>
	</tr>
	<tr height=25>
		<td align="center" class=tablebody1><br><!--#include file=z_down_query.asp--></td>
	</tr>
</table>
<br>
<%
call down_payuser()
call down_manage()
conndown.close
set conndown=nothing	
call activeonline()
call footer()%>
