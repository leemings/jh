<%dim rsdown,sqldown,flagdown,rsp,rss,downpath,upmod
dim daytop,weektop,alltop,listnum,editnum,hotnum,eventnum,newnum,gudinnum,vipshow,sonum,paymod,retime,showset
sqldown="select flag from [admin] where username='"&membername&"'"
set rsdown=conndown.execute(sqldown)
if not(rsdown.eof and rsdown.bof) then
	flagdown=rsdown("flag")
else
	flagdown=0
end if
set rsp=conndown.execute("select downpath from [notdown]")
downpath=rsp("downpath")
if right(downpath,1)<>"/" then
	downpath=downpath&"/"
end if
rsp.close
set rsp=nothing
set rss=conndown.execute("select * from [showpage]")
daytop=rss("daytop")
weektop=rss("weektop")
alltop=rss("alltop")
listnum=rss("listnum")
editnum=rss("editnum")
eventnum=rss("eventnum")
hotnum=rss("hotnum")
newnum=rss("newnum")
gudinnum=rss("gudinnum")
vipshow=rss("vipshow")
sonum=rss("sonum")
paymod=rss("paymod")
retime=rss("retime")
showset=split(rss("showset"),",")
rss.close

sub down_nav()%>
	<table class=tableborder1 cellspacing=1 width="<%=Forum_body(12)%>" align=center>
		<tr>
			<th valign=middle height=25 width="100%">|<%
				dim rs1
				set rs1=server.createobject("adodb.recordset")
				sql="select class,classid from Dclass"
				rs1.open sql,conndown,1,3
				do while not rs1.eof
					%>&nbsp;&nbsp;<a href="z_down_index.asp?classid=<%=rs1("classid")%>" onMouseOver='ShowMenu(classlist<%=rs1("classid")%>,120)'><font color="ffffff"><%=rs1("class")%></font></a>&nbsp;&nbsp;|<%
					rs1.movenext
				loop
				rs1.close
				set rs1=nothing
			%></th>
		</tr>
	</table>
	<br>
<%end sub

'-------------上传方式--------------------------
upmod=0   '0:高速无组件上传，1:lyfupload组件上传
'-----------------------------------------------

function downshow(sx)
	if sx="" or sx=1 or sx=0 or sx=" " then
		response.write "<font color=red>开放下载</font>"
	elseif sx=2 then
		response.write "<font color=red>会员下载</font>"
	elseif sx=3 then
		response.write "<font color=red>VIP下载</font>"
	elseif sx=4 then
		response.write "<font color=red>特约下载</font>"
	elseif sx=5 then
		response.write "<font color=red>付费下载</font>"
	end if
end function

function soft(sx)
	if sx=0 then
		response.write "<font color=red>普通软件</font>"
	elseif sx=1 then
		response.write "<font color=red>RealPlay播放</font>"
	elseif sx=2 then
		response.write "<font color=red>MediaPlay播放</font>"
	elseif sx=3 then
		response.write "<font color=red>FLASH动画</font>"
	end if
end function

function downtime(addt,downt)
	dim dt
	if downt=0 then
		downtime="无限期下载"
	else
		dt=downt-(date()-addt)
		if date()-addt<=downt then
			downtime="有效期为"&downt&"天；还剩 "&dt&" 天"
		else
			downtime="有效期为"&downt&"天；已过期"
		end if
	end if
end function

function daydownload(dh,dd)
	'dh当日下载次数
	'dd限制日下载次数
	dim last
	if dd<>"" and dd<>0 then
		last=dd-dh
		if dh=<dd then
			daydownload="每日下载限制为"&dd&"次；还有<font color=blue>"&last&"</font>次"
		else
			daydownload="每日下载限制为"&dd&"次；已截止"
		end if
	else
		daydownload="每日下载次数无限制"
	end if
end function 

function classnum(expression)
	dim rs_class,sql_class
	rs_class=conn.execute("select userclass from usertitle where usertitle='"&expression&"'")
	classnum=rs_class("userclass")
end function

sub down_manage()
	dim rsd,sqld,flagd
	sqld="select flag from [admin] where username='"&membername&"'"
	set rsd=conndown.execute(sqld)
	if not(rsd.eof and rsd.bof) then
		flagd=rsd("flag")%>
		<br>
		<table class=tableborder1 cellspacing=1 width="<%=Forum_body(12)%>" align=center>
			<tr>
				<th valign=middle height=25 width="100%">下载中心管理</th>
			</tr>
			<tr>
				<td class=tablebody1 align=center height=25><b><%=membername%></b> 您可以进行的管理项目包括：&nbsp;&nbsp;&nbsp;<%
					if flagd=1 then
						%>◇<a href="z_down_softadd.asp">添加上传软件</a>&nbsp;&nbsp;&nbsp;◇<a href="z_down_freeadd.asp">添加连接软件</a>&nbsp;&nbsp;&nbsp;◇<a href="z_down_adminedit.asp">修改删除软件</a>&nbsp;&nbsp;&nbsp;◇<a href="admin_index.asp">人员管理，栏目管理请进后台</a><%
					elseif flagd=2 then
						%>◇<a href="z_down_softadd.asp">添加上传软件</a>&nbsp;&nbsp;&nbsp;◇<a href="z_down_freeadd.asp">添加连接软件</a>&nbsp;&nbsp;&nbsp;◇<a href="z_down_adminedit.asp">修改删除软件</a><%
					elseif flagd=3 then
						%>◇<a href="z_down_softadd.asp">添加上传软件</a>&nbsp;&nbsp;&nbsp;◇<a href="z_down_freeadd.asp">添加连接软件</a><%
					elseif flagd=4 then
						%>◇<a href="z_down_freeadd.asp">添加连接软件</a><%
					end if
				%></td>
			</tr>
		</table>
	<%end if
end sub

'-------------------付费用户--------------------
sub down_payuser()
	dim rsp,sqlp
	sqlp="select * from [user] where username='"&membername&"' "
	set rsp=conndown.execute(sqlp)
	if not(rsp.eof and rsp.bof) then
		if rsp("apply")=2 then%>
			<br>
			<table class=tableborder1 cellspacing=1 width="<%=Forum_body(12)%>" align=center>
				<tr>
					<th valign=middle height=25 width="100%">付费用户控制面版</th>
				</tr>
				<tr>
					<td class=tablebody1 align=center height=30>您已经递交过申请了，您的申请暂未审批，请等待！</td>
				</tr>
			</table>
		<%elseif rsp("apply")=1 then%>
			<br>
			<table class=tableborder1 cellspacing=1 width="<%=Forum_body(12)%>" align=center>
				<tr>
					<th valign=middle height=25 width="100%">付费用户控制面版</th>
				</tr>
				<%if session("payuser")<>membername then%>
					<form action=z_down_paylogin.asp?action=chk method=post>
					<tr>
						<td class=tablebody1 align=center height=30><b><%=membername%></b> 请输入您的密码进入付费用户控制面版：　<input maxLength=20 name=password size=12 type=password>&nbsp;&nbsp;<input type=submit name=submit value="进 入"></td>
					</tr>
					</form>
				<%elseif session("payuser")=membername then%>
					<tr>
						<td class=tablebody1 align=center height=30><b><%=membername%></b> &nbsp;&nbsp;[<a href=z_down_paycontrol.asp>查看详细信息</a>]&nbsp;&nbsp;&nbsp;[<a href=z_down_paylogout.asp>退出</a>]</td>
					</tr>
				<%end if%>
			</table>
		<%end if
	elseif membername<>"" then%>
		<br>
		<table class=tableborder1 cellspacing=1 width="<%=Forum_body(12)%>" align=center>
			<tr>
				<th valign=middle height=25 width="100%">付费用户申请</th>
			</tr>
			<tr>
				<td class=tablebody1 align=center height=30><a href=z_down_applypayuser.asp><font color=red>[申请成为付费用户]</font></a></td>
			</tr>
		</table>
	<%end if
end sub

'--------------------付费地址显示
sub paybuy()
	dim rsu,foundsee,seenameid,seecount
	sql="select id from [user] where username='"&membername&"' and lockuser=1 and apply=1"
	set rsu=conndown.execute(sql)
	if rsu.eof and rsu.bof then
		userid=-1
	else
		userid=rsu("id")
	end if
	foundsee=false
	sql="select * from [download] where id="&sid
	set rsu=conndown.execute(sql)
	if rsu.eof and rsu.bof then
		response.write "此软件不存在！"
	else
		if lcase(trim(rsu("adduser")))=lcase(trim(membername)) or (flagdown>=1 and flagdown<=3) then
			foundsee=true
		else
			seenameid=split(rsu("buyuser"),",")
			seecount=ubound(seenameid)
			for i = 0 to seecount
				if strcomp(seenameid(i),userid)=0 then
					foundsee=true
					exit for
				end if
			next
		end if
		if foundsee then
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
			if lcase(trim(rsu("adduser")))=lcase(trim(membername)) then
				%><font color=<%=forum_body(8)%>>您是软件提供者无须付钱</font><br>[您可以下载此软件]<br><%
			elseif flagdown>=1 and flagdown<=3 then
				%><font color=<%=forum_body(8)%>>您是下载中心管理者无须付钱</font><br>[您可以下载此软件]<br><%
			else
				%><font color=<%=forum_body(8)%>>您已经付了 <b><%=rsu("point")%></b> 元现金</font><br>[您可以下载此软件]<br><%
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
		else
			if userid>0 then
				%><form action=z_down_paybuy.asp method=post><b>下载地址：</b><font color=red>下载此软件需要 <b><%=rsu("point")%></b> 元现金</font><br>&nbsp;&nbsp;<font color=gray>如果您想下载此软件，请输入您的付费中心密码，然后点击付费</font><br>&nbsp;&nbsp;<input maxLength=20 name=sid size=12 type=hidden value="<%=rsu("id")%>"><input maxLength=20 name=password size=12 type=password>&nbsp;&nbsp;<input type=submit name=submit value="付费"></form><%
			else
				%><b>下载地址：</b><font color=red>下载此软件需要 <b><%=rsu("point")%></b> 元现金</font><%               
			end if
		end if
	end if
end sub%>
