<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_down_conn.asp"-->
<!--#include file="z_down_function.asp" -->
<%dim menu
menu=request.querystring("menu")
stats="�����û�����"
call nav()
call head_var(0,0,"��������","z_down_default.asp")
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>��û���������ĸ����û������Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
	call dvbbs_error()
else
	if menu="" then
 		call main()
	elseif menu=1 then
  	call apply()
	elseif menu=2 then
  	call cancel()
	end if  	
end if
conndown.close
set conndown=nothing
call activeonline()
call footer()

sub main()
	sql="select * from [payconfig]"
	set rs=conndown.execute(sql)%>
	<table cellspacing=1 align=center class=tableborder1>
		<tr>
			<th valign=middle align=center height=25 width="100%" colspan="4">����˵��</th>
		</tr>
		<tr>
			<td valign=middle align=center height=25 width="100%" class=tablebody1 colspan="4"><%
				if paymod=1 then
					%><br><font color=red>������ϵͳ��ǰ���õ����Զ���Ӹ����û�ģʽ</font><br>�������ǰ���������㸶���û��������������Ϳ���ֱ�������Ϊ�������ĵĸ����û�<br><br><%
				else
					%><br><font color=red>������ϵͳ��ǰ���õ����˹���Ӹ����û�ģʽ</font><br>�������ǰ���������㸶���û��������������Ϳ��������Ϊ�������ĵĸ����û���<br>��������ᱻ�ύ������Ա������һ��ͨ���������ɳ�Ϊ�����û�<br><br><%
				end if
			%></td>
		</tr>
		<tr>
			<td valign=middle align=center height=25 width="100%" class=tablebody1 colspan="4">��ӭ<B><%=membername%></B>�����������ĸ����û�<br></td>
		</tr>
		<tr>
			<th valign=middle align=center height=25 width="25%"><b>����Ҫ����</b></th>
			<th valign=middle align=center width="25%">�����û�<b>����</b></th>
			<th valign=middle align=center width="25%"><b>����ǰ������</b></th>
			<th valign=middle align=center width="25%"><b>�Ƿ�����</b></th>
		</tr>
		<tr>
			<td valign=middle align=center height=25 width="25%" class=tablebody1>��Ǯ</td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("wealth")<=0 then
					response.write "-"
				else
					response.write rs("wealth")
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("wealth")<=0 then
					response.write "-"
				else
					response.write mymoney
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("wealth")<=0 then
					response.write "-"
				else
					response.write iiif(rs("wealth"),mymoney,"����","<font color=red>������</font>")
				end if
			%></td>
		</tr>
		<tr>
			<td valign=middle align=center height=25 width="25%" class=tablebody1>������</td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("article")<=0 then
					response.write "-"
				else
					response.write rs("article")
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("article")<=0 then
					response.write "-"
				else
					response.write myarticle
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("article")<=0 then
					response.write "-"
				else
					response.write iiif(rs("article"),myarticle,"����","<font color=red>������</font>")
				end if
			%></td>
		</tr>
		<tr>
			<td valign=middle align=center height=25 width="25%" class=tablebody1>����ֵ</td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("userep")<=0 then
					response.write "-"
				else
					response.write rs("userep")
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("userep")<=0 then
					response.write "-"
				else
					response.write myuserep
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("userep")<=0 then
					response.write "-"
				else
					response.write iiif(rs("userep"),myuserep,"����","<font color=red>������</font>")
				end if
			%></td>
		</tr>
		<tr>
			<td valign=middle align=center height=25 width="25%" class=tablebody1>����ֵ</td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("usercp")<=0 then
					response.write "-"
				else
					response.write rs("usercp")
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("usercp")<=0 then
					response.write "-"
				else
					response.write myusercp
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("usercp")<=0 then
					response.write "-"
				else
					response.write iiif(rs("usercp"),myusercp,"����","<font color=red>������</font>")
				end if
			%></td>
		</tr>
		<tr>
			<td valign=middle align=center height=25 width="25%" class=tablebody1>����ֵ</td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("power")<=0 then
					response.write "-"
				else
					response.write rs("power")
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("power")<=0 then
					response.write "-"
				else
					response.write mypower
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("power")<=0 then
					response.write "-"
				else
					response.write iiif(rs("power"),mypower,"����","<font color=red>������</font>")
				end if
			%></td>
		</tr>
		<tr>
			<td valign=middle align=center height=25 width="25%" class=tablebody1>����ֵ</td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("userscore")<=0 then
					response.write "-"
				else
					response.write rs("userscore")
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("userscore")<=0 then
					response.write "-"
				else
					response.write myuserscore
				end if
			%></td>
			<td valign=middle align=center width="25%" class=tablebody1><%
				if rs("userscore")<=0 then
					response.write "-"
				else
					response.write iiif(rs("userscore"),myuserscore,"����","<font color=red>������</font>")
				end if
			%></td>
		</tr>
		<tr>
			<td valign=middle align=center width="50%" colspan="4" class=tablebody1><%
				if rs("wealth")>mymoney or rs("article")>myarticle or rs("userep")>myuserep or rs("usercp")>myusercp or rs("power")>mypower or rs("userscore")>myuserscore then
					response.write "<br><font color="&forum_body(8)&">��������һ�����������㣬����ʱ���޷������Ϊ"&Forum_info(0)&"�������ĸ��ѻ�Ա</font><br>"
				else
					set rs=conndown.execute("select * from [user] where username='"&membername&"'")
					if rs.eof and rs.bof then
						response.write "<br><font color="&forum_body(8)&">�����������Ϊ�������ĸ��ѻ�Ա</font><br>"
						%><form name="form2" method="post" action="?menu=1"><INPUT type=submit value=���� name=submit></form><%
					else
						if rs("apply")=2 then
							response.write "<br><font color="&forum_body(8)&">�������뻹δ����</font><br>"
						elseif rs("apply")=1 then
							response.write "<br><font color="&forum_body(8)&">�������������ĵĸ��ѻ�Ա</font><br>"
						end if
					end if
				end if
			%></td>
		</tr>
	</table>
<%end sub

sub apply()
	dim rss,errword
	sql="select * from [payconfig]"
	set rs=conndown.execute(sql)
	set rss = server.CreateObject ("Adodb.recordset")
	sql="select * from [user] where username='"&membername&"'"
	rss.open sql,conndown,1,3
	if rs("wealth")>mymoney or rs("article")>myarticle or rs("userep")>myuserep or rs("usercp")>myusercp or rs("power")>mypower or rs("userscore")>myuserscore then
		errword="��������һ�����������㣬����ʱ���޷������Ϊ"&Forum_info(0)&"�������ĵĸ����û�"%>
		<table cellspacing=1 align=center class=tableborder1>
			<tr>
				<th valign=middle align=center height=25 width="100%">���븶���û�</th>
			</tr>
			<tr>
				<td valign=middle align=center height=25 width="100%" class=tablebody1><font color=red><%=errword%></font></td>
			</tr>
			<tr>
				<td valign=middle align=center height=25 width="100%" class=tablebody2><a href=z_down_default.asp>��������������ҳ</a></td>
			</tr>
		</table>
	<%elseif not(rss.eof and rss.bof) then
		errword="���ѵݽ�������������Ѿ��Ǹ��ѻ�Ա�ˣ��벻Ҫ�ظ�����"%>
		<table cellspacing=1 align=center class=tableborder1>
			<tr>
				<th valign=middle align=center height=25 width="100%">���븶���û�</th>
			</tr>
			<tr>
				<td valign=middle align=center height=25 width="100%" class=tablebody1><font color=<%=forum_body(8)%>><%=errword%></font></td>
			</tr>
			<tr>
				<td valign=middle align=center height=25 width="100%" class=tablebody2><a href=z_down_default.asp>��������������ҳ</a></td>
			</tr>
		</table>
	<%else
		rss.addnew
		rss("username")=membername
		rss("regtime")=date()
		rss("allpoint")=0
		rss("lockuser")=1
		if paymod=1 then
			rss("apply")=1 
		else
			rss("apply")=2 
		end if
		rss.update%>
		<table cellspacing=1 align=center class=tableborder1>
			<tr>
				<th valign=middle align=center height=25 width="100%">���븶���û�</th>
			</tr>
			<tr>
				<td valign=middle align=center height=25 width="100%" class=tablebody1><font color=<%=forum_body(8)%>><%
					if paymod=1 then
						response.write "���������Ѿ��ɹ�����ϲ����Ϊ"&Forum_info(0)&"�������ĵĸ����û�"
					else
						response.write "���������Ѿ��ݽ�����ȴ�����Ա����������л�������Ϊ"&Forum_info(0)&"�������ĵĸ����û�"
					end if
				%></font></td>
			</tr>
			<tr>
				<td valign=middle align=center height=25 width="100%" class=tablebody2><a href=z_down_default.asp>��������������ҳ</a></td>
			</tr>
		</table>
	<%end if
end sub%>
