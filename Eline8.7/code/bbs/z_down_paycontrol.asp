<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_down_conn.asp"-->
<!--#include file="z_down_function.asp" -->
<%
stats="�����û��������"
call nav()
call head_var(0,0,"��������","z_down_default.asp")
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>��û�н��븶���û���������Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
	call dvbbs_error()
elseif session("payuser")<>membername then
	Errmsg=Errmsg+"<br>"+"<li>��û�н��븶���û���������Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
	call dvbbs_error()
else
	sql="select * from [user] where username='"&membername&"'"
	set rs=conndown.execute(sql)%>
	<table border="0" cellspacing="1" cellpadding="3" align=center width="<%=forum_body(12)%>">
		<tr>
			<td width="27%">
				<table class=tableborder1 cellspacing=1 width="100%" align=center>
					<tr>
						<th valign=middle height=25 width="100%" colspan="2">�û�������Ϣ</th>
					</tr>
					<tr>
						<td class=tablebody1 align=right height=25 width="50%">�˺�״̬��&nbsp;</td>
						<td class=tablebody1 align=left>&nbsp;<%if rs("lockuser")=1 then%>����<%else%>����<%end if%></td>
					</tr>
					<tr>
						<td class=tablebody1 align=right height=25 width="50%">�û�����&nbsp;</td>
						<td class=tablebody1 align=left>&nbsp;<%=rs("username")%></td>
					</tr>
					<tr>
						<td class=tablebody1 align=right height=25 width="50%">ע��ʱ�䣺&nbsp;</td>
						<td class=tablebody1 align=left>&nbsp;<%=rs("regtime")%></td>
					</tr>
					<tr>
						<td class=tablebody1 align=right height=25 width="50%">�����ֽ�&nbsp;</td>
						<td class=tablebody1 align=left>&nbsp;<b><font color=<%=forum_body(8)%>><%=mymoney%></b></font></td>
					</tr>
					<tr>
						<td class=tablebody1 align=right height=25 width="50%">�ۻ������ֽ�&nbsp;</td>
						<td class=tablebody1 align=left>&nbsp;<b><font color=<%=forum_body(8)%>><%=rs("allpoint")%></b></font></td>
					</tr>
				</table>
			</td>
			<td width="73%" valign="top" align="left">
				<table class=tableborder1 cellspacing=1 width="100%" align=center>
					<tr>
						<th valign=middle height=25 width="100%">�û�������Ϣ</th>
					</tr>
					<tr>
						<td class=tablebody1 align=center>
							<table border="0" width="97%" cellspacing="1" align=center >
								<%sql="select top 10 * from [downevent] where username='"&membername&"' order by addtime desc"
								set rs=conndown.execute(sql)
								i=0
								if rs.eof and rs.bof then%>
									<tr>
										<td class=tablebody1 align=center height=25><b>û�����Ĺ�����Ϣ</b></td>
									</tr>
								<%else
									do while not rs.eof and i<10%>
										<tr>
											<td class=tablebody1 align=left height=25><font color=<%=forum_body(8)%>>��</font><%=rs("buydown")%>  <b>����<%=rs("addtime")%>����<font color=red><%=rs("buypoint")%>Ԫ</font></b></td>
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
						<th valign=middle height=25 width="100%">�û�������Ϣ</th>
					</tr>
					<tr>
						<td class=tablebody1 align=center>
							<table border="0" width="97%" cellspacing="1" align=center>
								<%sql="select top 10 * from [userevent] where username='"&membername&"' order by addtime desc"
								set rs=conndown.execute(sql)
								i=0
								if rs.eof and rs.bof then%>
									<tr>
										<td class=tablebody1 align=center height=25><b>û�����Ĳ�����Ϣ</b></td>
									</tr>
								<%else
									do while not rs.eof and i<10%>
										<tr>
											<td class=tablebody1 align=left height=25><font color=red>��</font><%=rs("userevent")%>  <b>����<%=rs("addtime")%>����</b>�������ˣ�<%=rs("conuser")%>��</td>     
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
