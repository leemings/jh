<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="z_plus_check.asp"-->
<%dim menu
menu=request.querystring("menu")
stats="VIP��Ա����"
call nav()
call head_var(2,0,"","")
if membername="" or not founduser then
	Errmsg=Errmsg+"<br>"+"<li>��û��VIP��Ա�����Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
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
call activeonline()
call footer() 

sub main()
	dim sql1,rs1,vipmod
	sql="select * from [vip]"
	set rs=conn.execute(sql)
	if not rs.eof then
		vipmod=rs("vipmod")
	end if
	sql1="select * from [user] where username='"&membername&"'"
	set rs1=conn.execute(sql1)
	if rs1.eof and rs1.bof then
		Errmsg=Errmsg+"<br>"+"<li>��������̳�ĺϷ��û���"
		exit sub
	end if%>
	<table cellspacing=1 align=center class=tableborder1 width="97%" height="228">
	<tr width="100%"><th valign=middle align=center height=21 width="100%" colspan="5">����˵��</td></tr>
	<tr width="100%">
	<td valign=middle align=left height=21 width="100%" class=tablebody1 colspan="5"><blockquote><%
		if vipmod=0 then
			%><br><font color=red>����Աϵͳ��ǰ���õ����Զ����VIP��Աģʽ</font><br>�������ǰ����������VIP��Ա�������������Ϳ���ֱ�������Ϊ<%=Forum_info(0)%>��VIP��Ա<%
		elseif vipmod=1 then
			%><br><font color=red>����Աϵͳ��ǰ���õ����˹����VIP��Աģʽ</font><br>�������ǰ����������VIP��Ա�������������Ϳ��������Ϊ<%=Forum_info(0)%>��VIP��Ա��<br>��������ᱻ�ύ������Ա������һ��ͨ���������ɳ�ΪVIP��Ա<%
		elseif vipmod=2 then
			%><br><font color=red>����Աϵͳ��ǰ���õ��ǰ��Զ����VIP��Աģʽ</font><br>�������ǰ����������VIP��Ա�״��������������Ϳ��������״γ�Ϊ<%=Forum_info(0)%>��VIP��Ա��<br>��������ᱻ�ύ������Ա������һ��ͨ���������ɳ�ΪVIP��Ա��<br>�������ǰ�����������ٴ��������������Ϳ���ֱ�������Ϊ<%=Forum_info(0)%>���ٴ�VIP��Ա<%
		end if
	%><br>1��VIP��Ա����Ч��Ϊ <b><%=rs("vipdate")%></b> ��<br>2����������Ч�ڿ쵽��ʱ�����ǻ���ǰ�ö�����ʽ֪ͨ��<br>3��лл����<%=Forum_info(0)%>��֧��<br></blockquote></td>
	</tr>
	<tr width="100%">
		<td valign=middle align=center height=21 width="100%" class=tablebody1 colspan="5">��ӭ<B><%=membername%></B>����VIP��Ա<br></td>
	</tr>
	<tr width="100%">
		<th valign=middle align=center height=24 width="20%"><b>�״�����Ҫ����</b></th>
		<th valign=middle align=center height=24 width="20%"><b>�״�VIP��Ա����</b></th>
		<th valign=middle align=center height=24 width="20%"><b>�ٴ�VIP��Ա����</b></th>
		<th valign=middle align=center height=24 width="20%"><b>����ǰ������</b></th>
		<th valign=middle align=center height=24 width="20%"><b>�Ƿ�����</b></th>
	</tr>
	<tr width="100%">
		<td valign=middle align=center height=21 width="20%" class=tablebody1>��Ǯ</td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if rs("wealth")<=0 then
				response.write "-"
			else
				response.write rs("wealth")
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if rs("wealth2")<=0 then
				response.write "-"
			else
				response.write rs("wealth2")
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if not isnull(rs1("vipdate")) then
				if rs("wealth2")<=0 then
					response.write "-"
				else
					response.write rs1("userwealth")-rs1("userwealth2")
				end if
			else
				if rs("wealth")<=0 then
					response.write "-"
				else
					response.write rs1("userwealth")
				end if
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if not isnull(rs1("vipdate")) then
				if rs("wealth2")<=0 then
					response.write "-"
				else
					response.write iiif(rs("wealth2"),rs1("userwealth")-rs1("userwealth2"),"����","<font color=red>������</font>")
				end if
			else	
				if rs("wealth")<=0 then
					response.write "-"
				else
					response.write iiif(rs("wealth"),rs1("userwealth"),"����","<font color=red>������</font>")
				end if
			end if
		%></td>
	</tr>
	<tr width="100%">
		<td valign=middle align=center height=21 width="20%" class=tablebody1>������</td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if rs("article")<=0 then
				response.write "-"
			else
				response.write rs("article")
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if rs("article2")<=0 then
				response.write "-"
			else
				response.write rs("article2")
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if not isnull(rs1("vipdate")) then
				if rs("article2")<=0 then
					response.write "-"
				else
					response.write rs1("article")-rs1("article2")
				end if
			else
				if rs("article")<=0 then
					response.write "-"
				else
					response.write rs1("article")
				end if
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if not isnull(rs1("vipdate")) then
				if rs("article2")<=0 then
					response.write "-"
				else
					response.write iiif(rs("article2"),rs1("article")-rs1("article2"),"����","<font color=red>������</font>")
				end if
			else	
				if rs("article")<=0 then
					response.write "-"
				else
					response.write iiif(rs("article"),rs1("article"),"����","<font color=red>������</font>")
				end if
			end if
		%></td>
	</tr>
	<tr width="100%">
		<td valign=middle align=center height=21 width="20%" class=tablebody1>����ֵ</td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if rs("userep")<=0 then
				response.write "-"
			else
				response.write rs("userep")
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if rs("userep2")<=0 then
				response.write "-"
			else
				response.write rs("userep2")
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if not isnull(rs1("vipdate")) then
				if rs("userep2")<=0 then
					response.write "-"
				else
					response.write rs1("userep")-rs1("userep2")
				end if
			else
				if rs("userep")<=0 then
					response.write "-"
				else
					response.write rs1("userep")
				end if
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if not isnull(rs1("vipdate")) then
				if rs("userep2")<=0 then
					response.write "-"
				else
					response.write iiif(rs("userep2"),rs1("userep")-rs1("userep2"),"����","<font color=red>������</font>")
				end if
			else	
				if rs("userep")<=0 then
					response.write "-"
				else
					response.write iiif(rs("userep"),rs1("userep"),"����","<font color=red>������</font>")
				end if
			end if
		%></td>
	</tr>
	<tr width="100%">
		<td valign=middle align=center height=21 width="20%" class=tablebody1>����ֵ</td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if rs("usercp")<=0 then
				response.write "-"
			else
				response.write rs("usercp")
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if rs("usercp2")<=0 then
				response.write "-"
			else
				response.write rs("usercp2")
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if not isnull(rs1("vipdate")) then
				if rs("usercp2")<=0 then
					response.write "-"
				else
					response.write rs1("usercp")-rs1("usercp2")
				end if
			else
				if rs("usercp")<=0 then
					response.write "-"
				else
					response.write rs1("usercp")
				end if
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if not isnull(rs1("vipdate")) then
				if rs("usercp2")<=0 then
					response.write "-"
				else
					response.write iiif(rs("usercp2"),rs1("usercp")-rs1("usercp2"),"����","<font color=red>������</font>")
				end if
			else	
				if rs("usercp")<=0 then
					response.write "-"
				else
					response.write iiif(rs("usercp"),rs1("usercp"),"����","<font color=red>������</font>")
				end if
			end if
		%></td>
	</tr>
	<tr width="100%">
		<td valign=middle align=center height=21 width="20%" class=tablebody1>����ֵ</td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if rs("power")<=0 then
				response.write "-"
			else
				response.write rs("power")
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if rs("power2")<=0 then
				response.write "-"
			else
				response.write rs("power2")
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if not isnull(rs1("vipdate")) then
				if rs("power2")<=0 then
					response.write "-"
				else
					response.write rs1("userpower")-rs1("userpower2")
				end if
			else
				if rs("power")<=0 then
					response.write "-"
				else
					response.write rs1("userpower")
				end if
			end if
		%></td>
		<td valign=middle align=center height=21 width="20%" class=tablebody1><%
			if not isnull(rs1("vipdate")) then
				if rs("power2")<=0 then
					response.write "-"
				else
					response.write iiif(rs("power2"),rs1("userpower")-rs1("userpower2"),"����","<font color=red>������</font>")
				end if
			else	
				if rs("power")<=0 then
					response.write "-"
				else
					response.write iiif(rs("power"),rs1("userpower"),"����","<font color=red>������</font>")
				end if
			end if
		%></td>
	</tr>
	<tr width="100%">
	<td valign=middle align=center width="25%" class=tablebody1><a href="toplist.asp?orders=9"><font color="#FF0000">��ѯVIP��Ա</font></a></td>
	<td valign=middle align=center width="50%" colspan="3" class=tablebody1><%
		if rs1("vip")=2 then
			response.write "<br><font color=red>�������ƣ����ܳ�Ϊ"&Forum_info(0)&"VIP��Ա</font><br><br>"
		elseif rs1("vip")=3 then
			response.write "<br><font color=red>���Ѿ���VIP��Ա��������VIPȨ���Ѿ����������벻Ҫ�ظ����룡</font><br><br>"
		elseif rs1("vip")=4 then
			response.write "<br><font color=red>���Ѿ��ݽ������룬��δ�õ���׼���벻Ҫ�ظ����룡</font><br><br>"
		elseif rs1("vip")=1 then
			response.write "<br><font color=red>���Ѿ���"&Forum_info(0)&"VIP��Ա</font><br>"
			response.write "<br><font color=red>ע�����Ļ�Ա�ʸ�</font><br>"
			%><form name="form1" method="post" action="z_vip.asp?menu=2"><INPUT name=username size="20" value="<%=membername%>" type=hidden><INPUT type=submit value=ע�� name=submit></form><%
		else
			if not isnull(rs1("vipdate")) then
				if (rs("wealth2")>rs1("userwealth")-rs1("userwealth2") and rs("wealth2")>0) or (rs("article2")>rs1("article")-rs1("article2") and rs("article2")>0) or (rs("userep2")>rs1("userep")-rs1("userep2") and rs("userep2")>0) or (rs("usercp2")>rs1("usercp")-rs1("usercp2") and rs("usercp2")>0) or (rs("power2")>rs1("userpower")-rs1("userpower2") and rs("power2")>0) then
					response.write "<br><br><font color=red>��������һ�����������㣬����ʱ���޷��ٴ������Ϊ"&Forum_info(0)&"VIP��Ա</font><br><br>"
				else
					response.write "<br><font color=red>����������VIP��Ա</font><br>"
					%><form name="form2" method="post" action="z_vip.asp?menu=1"><INPUT name=username size="20" value="<%=membername%>" type=hidden><INPUT type=submit value=���� name=submit></form><%
				end if
			else
				if (rs("wealth")>rs1("userwealth") and rs("wealth")>0) or (rs("article")>rs1("article") and rs("article")>0) or (rs("userep")>rs1("userep") and rs("userep")>0) or (rs("usercp")>rs1("usercp") and rs("usercp")>0) or (rs("power")>rs1("userpower") and rs("power")>0) then
					response.write "<br><br><font color=red>��������һ�����������㣬����ʱ���޷��״������Ϊ"&Forum_info(0)&"VIP��Ա</font><br><br>"
				else
					response.write "<br><font color=red>����������VIP��Ա</font><br>"
					%><form name="form2" method="post" action="z_vip.asp?menu=1"><INPUT name=username size="20" value="<%=membername%>" type=hidden><INPUT type=submit value=���� name=submit></form><%
				end if
			end if
		end if
	%></td>
	<td valign=middle align=center width="25%" class=tablebody1><a href="toplist.asp?orders=10"><font color="#FF0000">�ȴ�VIP������Ա</font></a><br><br><%
		if rs1("vip")=1 then
			response.write "���ϴγ�ΪVIP��Ա��ʱ��Ϊ��<br>"&formatdatetime(rs1("vipdate"),2)&"<br>"
			response.write "��������Ч�ڻ��У�<b>"&cint(rs("vipdate")-datediff("d",rs1("vipdate"),now()))&"</b>��"
		end if
	%></td>
	</tr>
	</table>
	<%rs.close
	Set rs = Nothing
	rs1.close
	Set rs1 = Nothing
end sub

sub apply()
	dim boarduser
	dim boarduser_1
	dim userlen
	dim updateinfo
	dim rs1,sql1,rs,sql,err
	sql="select * from [vip]"
	set rs=conn.execute(sql)
	sql1="select * from [user] where username='"&membername&"'"
	set rs1=conn.execute(sql1)
	if rs1("vip")=1 then
		Errmsg=Errmsg+"<br>"+"<li>���Ѿ���"&Forum_info(0)&"��VIP��Ա���벻Ҫ�ظ����룡"
		call dvbbs_error()
	elseif rs1("vip")=2 then
		Errmsg=Errmsg+"<br>"+"<li>�������ƣ����ܳ�Ϊ"&Forum_info(0)&"VIP��Ա��"
		call dvbbs_error()
	elseif rs1("vip")=3 then
		Errmsg=Errmsg+"<br>"+"<li>���Ѿ���VIP��Ա��������VIPȨ���Ѿ����������벻Ҫ�ظ����룡"
		call dvbbs_error()
	elseif rs1("vip")=4 then
		Errmsg=Errmsg+"<br>"+"<li>���Ѿ��ݽ������룬��δ�õ���׼���벻Ҫ�ظ����룡"
		call dvbbs_error()
	else
		dim sucword
		sucword=""
		if not isnull(rs1("vipdate")) then
			if (rs("wealth2")>rs1("userwealth")-rs1("userwealth2") and rs("wealth2")>0) or (rs("article2")>rs1("article")-rs1("article2") and rs("article2")>0) or (rs("userep2")>rs1("userep")-rs1("userep2") and rs("userep2")>0) or (rs("usercp2")>rs1("usercp")-rs1("usercp2") and rs("usercp2")>0) or (rs("power2")>rs1("userpower")-rs1("userpower2") and rs("power2")>0) then
				Errmsg=Errmsg+"<br>"+"<li>��������һ�����������㣬����ʱ���޷��ٴ������Ϊ"&Forum_info(0)&"��VIP��Ա"
				call dvbbs_error()
			else
				if rs("vipmod")=0 or rs("vipmod")=2 then
					conn.execute("update [user] set vip=1,vipdate=now(),userwealth2=userwealth,userep2=userep,usercp2=usercp,userpower2=userpower,article2=article where username='"&membername&"'")
					sucword="������ɹ�����ϲ���ٴγ�Ϊ"&Forum_info(0)&"��VIP��Ա"
				elseif rs("vipmod")=1 then
					conn.execute("update [user] set vip=4 where username='"&membername&"'")
					sucword="���������Ѿ��ύ����ȴ�����Ա��������"
				end if
			end if
		else
			if (rs("wealth")>rs1("userwealth") and rs("wealth")>0) or (rs("article")>rs1("article") and rs("article")>0) or (rs("userep")>rs1("userep") and rs("userep")>0) or (rs("usercp")>rs1("usercp") and rs("usercp")>0) or (rs("power")>rs1("userpower") and rs("power")>0) then
				Errmsg=Errmsg+"<br>"+"<li>��������һ�����������㣬����ʱ���޷��״������Ϊ"&Forum_info(0)&"��VIP��Ա"
				call dvbbs_error()
			else
				if rs("vipmod")=0 then
					conn.execute("update [user] set vip=1,vipdate=now(),userwealth2=userwealth,userep2=userep,usercp2=usercp,userpower2=userpower,article2=article where username='"&membername&"'")
					sucword="������ɹ�����ϲ���״γ�Ϊ"&Forum_info(0)&"��VIP��Ա"
				elseif rs("vipmod")=1 or rs("vipmod")=2 then
					conn.execute("update [user] set vip=4 where username='"&membername&"'")
					sucword="���������Ѿ��ύ����ȴ�����Ա��������"
				end if
			end if
		end if
		if sucword<>"" then
			%><table cellspacing=1 align=center class=tableborder1 width="97%" >
				<tr width="100%"><th valign=middle align=center height=21 width="100%">�״�����VIP��Ա</tr>
				<tr width="100%">
					<td valign=middle align=center height=21 width="100%" class=tablebody1><br><font color=red><%=sucword%></font><br><br><a href=index.asp>������ҳ</a><br></td>
				</tr>
			</table><%
		end if
	end if
	rs1.close
	set rs1=nothing
	rs.close
	set rs=nothing
end sub

sub cancel()
	dim boarduser
	dim boarduser_1
	dim userlen
	dim updateinfo
	dim rs1,sql1,rs,sql
	sql="select * from [vip]"
	set rs=conn.execute(sql)
	conn.execute("update [user] set vip=0 where username='"&membername&"'")
	%><table cellspacing=1 align=center class=tableborder1 width="97%" >
		<tr width="100%"><th valign=middle align=center height=21 width="100%">ע��VIP��Ա</tr>
		<tr width="100%">
			<td valign=middle align=center height=21 width="100%" class=tablebody1><br><font color=red>��ע���ɹ�����ӭ����Ϊ<%=Forum_info(0)%>����ͨ��Ա</font><br><br><a href=index.asp>������ҳ</a><br></td>
		</tr>
		</table><%
	rs.close
	set rs=nothing
end sub%>



