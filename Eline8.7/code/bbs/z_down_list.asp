<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_down_conn.asp"-->
<!--#include file="z_down_function.asp" -->
<!--#include file="inc/ubbcode.asp" -->
<!--#include file="md5.asp" -->
<%
dim downpwd,rs_nodown,classname,Nclassname,sid,play,pimg
set rs_nodown=conndown.execute("select * from notdown")
downpwd=rs_nodown("nodown")
rs_nodown.close
set rs_nodown=nothing
stats="下载软件"
if request("down")<>"" then
	call downsoft()
	response.end
end if
call nav()
call head_var(0,0,"下载中心","z_Down_default.asp")
if request("id")="" then
	response.write "您没有选择相关软件，请返回"
	response.end
end if
set rs=server.createobject("adodb.recordset")
sql="select Dclass.class,DNclass.Nclass from download,Dclass,DNclass where download.classid=Dclass.classid and download.Nclassid=DNclass.Nclassid and download.ID="&request("id")
rs.open sql,conndown,1,1
if not rs.eof then
	classname=rs("class")
	Nclassname=rs("Nclass")
end if
rs.close%>
<script language="JavaScript">
function Showvote(id) {
	var filename="z_down_vote.asp?id="+id;
	window.open(filename,"显示窗口","scrollbars=yes,width=350,height=300");
}
</SCRIPT>
<!-- #include file="z_down_forum_js.asp" -->
<%call down_nav()%>
<table align=center border="0" width="<%=Forum_body(12)%>">
	<tr>
		<td width="25%" valign="top" align="left"><%
			if showset(10)=1 then
				%><!--#include file="z_down_allhot.asp"--><br><%
			end if
			if showset(11)=1 then
				%><!--#include file="z_down_hot.asp"--><br><%
			end if
			if showset(12)=1 then
				%><!--#include file="z_down_day.asp"--><br><%
			end if
			if showset(13)=1 then
				%><!--#include file="z_down_week.asp"--><br><%
			end if
			if showset(14)=1 then
				%><!--#include file="z_down_all.asp"--><br><%
			end if
		%></td>
		<td width="75%" valign="top" align="left">
			<table cellspacing=1 width="100%" align=center  style="word-break:break-all;" bgcolor=<%=Forum_Body(27)%>>
				<%sid=request("id") 
				sql="select * from download where id="&request("id")   
				rs.open sql,conndown,1,1
				sid=rs("id")%>   
				<tr>
					<th width="100%" height="24">□- <a href="z_down_default.asp"><font  color="ffffff">下载分类</font></a> >> <a href='z_down_index.asp?classid=<%=rs("classid")%>'><font color="ffffff"><%=classname%></font></a> >> <a  href='z_down_index.asp?classid=<%=rs("classid")%>&Nclassid=<%=rs("Nclassid")%>'><font  color="ffffff"><%=Nclassname%></font></a> >> <%=rs("showname")%></th>
				</tr> 
				<tr>
					<td width="100%" class=tablebody1>  
						<table width="100%" align="center" class=tablebody1>
							<tr> 
								<td colspan="2" class=tablebody1 align="center"><b><%=rs("showname")%></b></td>
							</tr>
							<tr> 
								<td colspan="2" class=tablebody1>&nbsp;</td>
							</tr>
							<tr> 
								<td class=tablebody1 width="50%"><b>软件名称：</b><b><%=rs("showname")%></b></td>
								<td class=tablebody1 width="50%"><b>软件属性：</b><%=soft(rs("softsx"))%></td>
							</tr>
							<tr> 
								<td colspan="2" class=tablebody1>&nbsp;</td>
							</tr>
							<tr> 
								<td class=tablebody1 width="50%"><b>下载方式：</b><%=downshow(rs("downshow"))%></td>
								<td class=tablebody1 width="50%"><b>授权方式：</b><%=rs("orders")%></td>
							</tr>
							<tr> 
								<td colspan="2" class=tablebody1>&nbsp;</td>
							</tr>
							<tr> 
								<td width="50%" class=tablebody1><b>软件类型：</b><%=Nclassname%></td>
								<td width="50%" class=tablebody1><b>软件评价：</b><%for i=1 to rs("hot")%><img  src="images/img_down/friend.gif"><%next%></td>
							</tr>
							<tr> 
								<td width="50%" class=tablebody1>&nbsp;</td>
								<td width="50%" class=tablebody1>&nbsp;</td>
							</tr>
							<tr> 
								<td width="50%" class=tablebody1><b>软件大小：</b><%=rs("size")%></td>
								<td width="50%" class=tablebody1><b>相关链接：</b><a href="<%=rs("fromurl")%>">主页</a></td>              
							</tr>
							<tr> 
								<td colspan=2 class=tablebody1>&nbsp;</td>
							</tr>
							<tr> 
								<td width="50%" class=tablebody1><b>上传时间：</b><%=rs("dateandtime")%></td>
								<td width="50%" class=tablebody1><b>本日下载：</b><%=rs("dayhits")%>&nbsp;&nbsp;本周： <%=rs("weekhits")%>&nbsp;&nbsp;总计：<%=rs("hits")%></td>             
							</tr>
							<tr> 
								<td colspan=2 class=tablebody1>&nbsp;</td>
							</tr>
							<tr> 
								<td class=tablebody1 width="50%"><b>运行环境：</b><%=rs("system")%></td>
								<td class=tablebody1 width="50%"><b>本软件相关图片：</b></td>
							</tr>
							<tr> 
								<td class=tablebody1 width="50%">&nbsp;</td>
								<td class=tablebody1 width="50%" rowspan="2" valign="top" align="left"><%
									if rs("pic")<>"" and rs("pic")<>" " then
										%> <a href=z_down_showpic.asp?id=<%=rs("id")%> target=_blank><img src="<%=rs("pic")%>" border="0" width="200" height="150" alt=点击查看详细图片></a><%
									else
										%>此软件无相关图片<%
									end if
								%></td>
							</tr>
							<tr> 
								<SCRIPT language="JavaScript">     
								function windowOpen(loadpos) {
									controlWindow=window.open(loadpos,"surveywin","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,width=470,height=400,status=yes,resizable=yes");     
								}    
								</SCRIPT>
								<td class=tablebody1 width="50%" valign="top" align="left"><%
									if (rs("usetime")=0 or (date()-rs("dateandtime")<=rs("usetime"))) and (rs("daydown")=0 or rs("daydown")=NULL or (rs("dayhits")<=rs("daydown"))) then
										response.write "<b>下载有效期：</b>"&downtime(rs("dateandtime"),rs("usetime"))&"<br><br>"
										response.write "<b>每日下载次数限制：</b>"&daydownload(rs("dayhits"),rs("daydown"))&"<br><br>"     
										if rs("downshow")=1 then
											call showurl()
										end if
										dim sql_u,rs_u
										if vipshow=2 then
											sql_u="select vip,userclass from [user] where username='"&membername&"'"
											set rs_u=conn.execute(sql_u)
										elseif vipshow=1 then
											sql_u="select userclass from [user] where username='"&membername&"'"
											set rs_u=conn.execute(sql_u)
										end if
										if rs("downshow")=2 then
											if rs_u.eof and rs_u.bof then
												response.write "<b>下载地址：</b><font color=red>此软件只对本论坛会员开放下载</font>"
											else
												call showurl()
											end if
										end if
										if rs("downshow")=3 then
											if rs_u.eof and rs_u.bof then
												response.write "<b>下载地址：</b><font color=red>此软件只对本论坛会员开放下载</font>"
											else 
												if rs_u("vip")=1 then
													call showurl()
												else
													response.write "<b>下载地址：</b><font color=red>此软件只对本论坛VIP会员开放下载</font>"
												end if
											end if
										end if   
										if rs("downshow")=4 then
											if rs_u.eof and rs_u.bof then
												response.write "<b>下载地址：</b><font color=red>此软件只对本论坛会员开放下载</font>"
											else 
												if flagdown<=0 or flagdown>4 then
													response.write "<b>下载地址：</b><font color=red>此软件只对本论坛下载系统管理员开放</font>"
												else
													call showurl()
												end if
											end if
										end if
										if rs("downshow")=5 then
											call paybuy()
										end if
										response.write "<br><br>"
									else
										response.write "<b>下载有效期：</b><font color=red>"&downtime(rs("dateandtime"),rs("usetime"))&"</font><br><br>"
										response.write "<b>每日下载次数限制：</b><font color=red>"&daydownload(rs("dayhits"),rs("daydown"))&"</font><br><br>"
									end if
								%></td>
							</tr>
							<tr height=25> 
								<td colspan="2" class=tablebody1><b>软件简介：</b></td>
							</tr>
							<tr> 
								<td colspan="2" class=tablebody1><blockquote><%=DvBCode("[MARK]"&rs("note")&"[/MARK]",5,3)%></blockquote></td>
							</tr>
							<tr height=25> 
								<td colspan=2 class=tablebody1 align=center><b>提供人：</b><font color=red><%=forum_info(0)%></font>  <b><%=rs("adduser")%></b></td>
							</tr>
							<tr height=25> 
								<td colspan=2 class=tablebody1></td>
							</tr>
							<tr height=25> 
								<td colspan="2" class=tablebody1>&nbsp;&nbsp;<font color=red>＊</font> 推荐使用网络蚂蚁下载本站软件，如果您发现链接错误或其它问题，请写信通知我们。<br>&nbsp;&nbsp;<font color=red>＊</font> 如果您链接本站软件，请注明来自<%=forum_info(0)%>，谢谢您的支持！<br>&nbsp;&nbsp;<font color=red>＊</font> 欢迎广大作者给我们提供软件以及使用说明。</td>
							</tr>
						</table>
					</td>
				</tr>   
				<tr>
					<td class=tablebody1 align="right" ><%call endshow()%></td>
				</tr>   
			</table>
		</td>                           
	</tr>                         
</table>   
<%call down_payuser()
call down_manage()
conndown.close
set conndown=nothing
call activeonline()
call footer()

sub showurl()
	if rs("softsx")=1 then     
		play="z_down_play.asp?action=rm"     
		pimg="images/img_down/rm.gif"     
	elseif rs("softsx")=2 then    
		play="z_down_play.asp?action=mpeg"     
		pimg="images/img_down/mpeg.gif"
	elseif rs("softsx")=3 then
		play="z_down_play.asp?action=swf"     
		pimg="images/img_down/swf.gif"
	elseif rs("softsx")=4 then
		play="z_down_play.asp?action=mp3"     
		pimg="images/img_down/mp3.gif" 
	end if
	if rs("softsx")>0 then 
		%><b>在线播放：</b><%
		if rs("filename")<>"" then
			%><A href="#" onClick="windowOpen('<%=play%>&id=<%=rs("id")%>&downid=1')"><IMG src="<%=pimg%>" BORDER=0 alt=在线播放地址一></A><%
		end if
		if rs("filename1")<>"" then
			%><A href="#" onClick="windowOpen('<%=play%>&id=<%=rs("id")%>&downid=2')"><IMG src="<%=pimg%>" BORDER=0 alt=在线播放地址二></A><%
		end if 
		if rs("filename2")<>"" then
			%><A href="#" onClick="windowOpen('<%=play%>&id=<%=rs("id")%>&downid=3')"><IMG src="<%=pimg%>" BORDER=0 alt=在线播放地址三></A><%
		end if
		%><br><br><%
	end if
	if lcase(left(rs("filename"),6))="ftp://" then
		%><b>下载地址1：</b><a href="<%=rs("filename")%>" ><IMG SRC="images/img_down/download.gif" BORDER=0 alt=下载地址一></a><%
	else
		if rs("ldown")=1 and instr(rs("filename"),"$")>0 then
			%><b>下载地址1：</b><a href="z_down_list.asp?id=<%=rs("id")%>&down=1&dpwd=<%=md5(downpwd&left(split(rs("filename"),"$")(1),6))%>" ><IMG SRC="images/img_down/download.gif" BORDER=0 alt=下载地址一></a><%
		else
			%><b>下载地址1：</b><a href="z_down_list.asp?id=<%=rs("id")%>&down=1&dpwd=<%=md5(downpwd)%>" ><IMG SRC="images/img_down/download.gif" BORDER=0 alt=下载地址一></a><%
		end if
	end if
	if rs("filename1")<>"" then
		if lcase(left(rs("filename1"),6))="ftp://" then
			%><br><b>下载地址2：</b><a href="<%=rs("filename1")%>" ><IMG SRC="images/img_down/download.gif" BORDER=0 alt=下载地址二></a><%
		else
			%><br><b>下载地址2：</b><a href="z_down_list.asp?id=<%=rs("id")%>&down=2&dpwd=<%=md5(downpwd)%>" ><IMG SRC="images/img_down/download.gif" BORDER=0 alt=下载地址二></a><%
		end if
	end if
	if rs("filename2")<>"" then
		if lcase(left(rs("filename2"),6))="ftp://" then
			%><br><b>下载地址3：</b><a href="<%=rs("filename2")%>" ><IMG SRC="images/img_down/download.gif" BORDER=0 alt=下载地址三></a><%
		else
			%><br><b>下载地址3：</b><a href="z_down_list.asp?id=<%=rs("id")%>&down=3&dpwd=<%=md5(downpwd)%>" ><IMG SRC="images/img_down/download.gif" BORDER=0 alt=下载地址三></a><%
		end if
	end if
	response.write "<br>"
end sub

sub endshow()%>
	<table align=center border="0" width="100%">
		<tr>
			<td class=tablebody1 align="left" width="60%"><a href="z_down_pag.asp?id=<%=sid%>"><img src="<%=Forum_info(7)&Forum_TopicPic(3)%>" border=0 alt=把本帖打包邮递></a>&nbsp;<a href="z_down_favadd.asp?id=<%=sid%>"><IMG SRC="<%=Forum_info(7)&Forum_TopicPic(4)%>" BORDER=0 alt=把本帖加入论坛收藏夹></a>&nbsp;<a href="z_down_sendpage.asp?id=<%=sid%>"><img src="<%=Forum_info(7)&Forum_TopicPic(5)%>" border=0 alt=发送本帖给朋友></a>&nbsp;<a href=#><span style="CURSOR: hand" onClick="window.external.AddFavorite('<%=Forum_info(1)%>/z_down_list.asp?id=<%=rs("id")%>', ' <%=Forum_info(0)%> - <%=htmlencode(replace(rs("showname"),"&#39;",""))%>')"><IMG SRC="<%=Forum_info(7)&Forum_TopicPic(6)%>" BORDER=0 width=15 height=15 alt=把本帖加入IE收藏夹></a>&nbsp;</td>
			<td class=tablebody1 align="right" width="40%">【<a href="javascript:Showvote('<%=rs("id")%>')"><font color="#FF0000">评论投票</font></a>】</td>
		</tr>
	</table>
<%end sub

sub downsoft()
	dim rdown,sqldown,downl,dpwd,ldown,URL
	Set rdown = Server.CreateObject("ADODB.Recordset")
	sqldown="SELECT * FROM download where id="&request("id")
	rdown.OPEN sqldown, Conndown,1,1
	downl=request("down")
	dpwd=request("dpwd")
	ldown=rdown("ldown")
	if downl>"3" then
		response.write "没有该下载地址"
		response.end
	elseif (rdown("ldown")<>1 or cint(downl)<>1) and dpwd<>md5(downpwd) then
		response.write "无权下载"
		response.end
	elseif rdown("ldown")=1 then
		if instr(rdown("filename"),"$")>0 then
			if dpwd<>md5(downpwd&left(split(rdown("filename"),"$")(1),6)) then
				response.write "无权下载"
				response.end
			end if
		else
			if dpwd<>md5(downpwd) then
				response.write "无权下载"
				response.end
			end if
		end if
		if downl="1" then
			if ldown=1 then
				URL=downpath&rdown("filename")
			else
				URL=rdown("filename")
			end if
		elseif downl="2" then
			URL=rdown("filename1")
		elseif downl="3" then
			URL=rdown("filename2")
		end if
	else
		if downl="1" then
			if ldown=1 then
				URL=downpath&rdown("filename")
			else
				URL=rdown("filename")
			end if
		elseif downl="2" then
			URL=rdown("filename1")
		elseif downl="3" then
			URL=rdown("filename2")
		end if
	end if
	rdown.close
	'更新每周每日数据
	dim LastHits
	sql="select lasthits from download where ID="&request("id")
	rdown.open sql,conndown,1,1
	if not rdown.eof then
		lasthits=rdown("lasthits")
	end if
	rdown.close
	dim tdate
	tdate=year(Now()) & "-" & month(Now()) & "-" & day(Now())
	if trim(lasthits)=trim(tdate) then
		sql="update download set dayhits=dayhits+1 where id="&request("id")
		conndown.Execute(sql)
	else
		sql="update download set dayhits=1 where id="&request("id")
		conndown.Execute(sql)
	end if
	sql="update download set hits=hits+1,lasthits='"&tdate&"' where ID="&request("id")
	conndown.Execute(sql)
	dim p_year,p_month,p_day,period_time
	p_year=CInt(year(Now()))-CInt(year(lasthits))
	p_month=CInt(month(Now()))-CInt(month(lasthits))
	p_day=CInt(day(Now()))-CInt(day(lasthits))
	period_time=((p_year*12+p_month)*30+p_day)
	if clng(period_time)<=7 then
		sql="update download set weekhits=weekhits+1 where id="&request("id")
		conndown.Execute(sql)
	else
		sql="update download set weekhits=1 where id="&request("id")
		conndown.Execute(sql)
	end if
	set rdown=nothing
	if left(url,1)<>"/" then
		response.redirect(url)
	else
		response.redirect "http://"&request.servervariables("SERVER_NAME")&"/"&url
	end if
end sub%>