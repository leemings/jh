<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_down_conn.asp"-->
<!--#include file="z_down_function.asp" -->
<%
stats="付费用户控制面版"
call nav()
call head_var(0,0,"下载中心","z_down_default.asp")
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>您没有进入付费用户控制面版的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
	call dvbbs_error()
elseif session("payuser")<>membername then
	Errmsg=Errmsg+"<br>"+"<li>您没有进入付费用户控制面版的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
	call dvbbs_error()
else
	sql="select * from [user] where username='"&membername&"'"
	set rs=conndown.execute(sql)%>
	<table border="0" cellspacing="1" cellpadding="3" align=center width="<%=forum_body(12)%>">
		<tr>
			<td width="27%">
				<table class=tableborder1 cellspacing=1 width="100%" align=center>
					<tr>
						<th valign=middle height=25 width="100%" colspan="2">用户基本信息</th>
					</tr>
					<tr>
						<td class=tablebody1 align=right height=25 width="50%">账号状态：&nbsp;</td>
						<td class=tablebody1 align=left>&nbsp;<%if rs("lockuser")=1 then%>正常<%else%>锁定<%end if%></td>
					</tr>
					<tr>
						<td class=tablebody1 align=right height=25 width="50%">用户名：&nbsp;</td>
						<td class=tablebody1 align=left>&nbsp;<%=rs("username")%></td>
					</tr>
					<tr>
						<td class=tablebody1 align=right height=25 width="50%">注册时间：&nbsp;</td>
						<td class=tablebody1 align=left>&nbsp;<%=rs("regtime")%></td>
					</tr>
					<tr>
						<td class=tablebody1 align=right height=25 width="50%">现有现金：&nbsp;</td>
						<td class=tablebody1 align=left>&nbsp;<b><font color=<%=forum_body(8)%>><%=mymoney%></b></font></td>
					</tr>
					<tr>
						<td class=tablebody1 align=right height=25 width="50%">累积花费现金：&nbsp;</td>
						<td class=tablebody1 align=left>&nbsp;<b><font color=<%=forum_body(8)%>><%=rs("allpoint")%></b></font></td>
					</tr>
				</table>
			</td>
			<td width="73%" valign="top" align="left">
				<table class=tableborder1 cellspacing=1 width="100%" align=center>
					<tr>
						<th valign=middle height=25 width="100%">用户下载信息</th>
					</tr>
					<tr>
						<td class=tablebody1 align=center>
							<table border="0" width="97%" cellspacing="1" align=center >
								<%sql="select top 10 * from [downevent] where username='"&membername&"' order by addtime desc"
								set rs=conndown.execute(sql)
								i=0
								if rs.eof and rs.bof then%>
									<tr>
										<td class=tablebody1 align=center height=25><b>没有您的购买信息</b></td>
									</tr>
								<%else
									do while not rs.eof and i<10%>
										<tr>
											<td class=tablebody1 align=left height=25><font color=<%=forum_body(8)%>>☆</font><%=rs("buydown")%>  <b>——<%=rs("addtime")%>——<font color=red><%=rs("buypoint")%>元</font></b></td>
										</tr>
										<%rs.movenext
										i=i+1
									loop
								end if
								rs.close%>
							</table>
						</td>
					</tr>
				</table>
				<br>
				<table class=tableborder1 cellspacing=1 width="100%" align=center>
					<tr>
						<th valign=middle height=25 width="100%">用户操作信息</th>
					</tr>
					<tr>
						<td class=tablebody1 align=center>
							<table border="0" width="97%" cellspacing="1" align=center>
								<%sql="select top 10 * from [userevent] where username='"&membername&"' order by addtime desc"
								set rs=conndown.execute(sql)
								i=0
								if rs.eof and rs.bof then%>
									<tr>
										<td class=tablebody1 align=center height=25><b>没有您的操作信息</b></td>
									</tr>
								<%else
									do while not rs.eof and i<10%>
										<tr>
											<td class=tablebody1 align=left height=25><font color=red>☆</font><%=rs("userevent")%>  <b>——<%=rs("addtime")%>——</b>（操作人：<%=rs("conuser")%>）</td>     
										</tr>
										<%rs.movenext
										i=i+1
									loop
								end if
								rs.close%>
							</table>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
<%end if
conndown.close
set conndown=nothing
call activeonline()
call footer()%>
