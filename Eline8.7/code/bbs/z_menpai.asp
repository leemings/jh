<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="z_menpaiconfig.asp" -->
<!--#include file="z_plus_check.asp"-->
<%
stats="��̳����"
call nav()
call head_var(2,0,"","")
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>��û�н�����̳���ɵ�Ȩ�ޣ����ȵ�¼����ͬ����Ա��ϵ��"
	call dvbbs_error()
else 
	call main()
end if
call activeonline()
call footer()
'============================================================
sub main()
%>
<table cellspacing=1 align=center border="0" class=tableborder1>
<tr><th valign=middle align=center height=24 width=<%=Forum_body(12)%> colspan="3" >����˵��</tr>
<tr>	
	<td valign=middle class=tablebody2 width="5%"></td>
	<td valign=middle class=tablebody1 width="90%">
	<ui>
		<li><u>��������</u>�������������˳�
		<li>ֻ��������Ҫ�����<u>��������</u>�������ſ��Լ��룬����ʱҪ����һ���Ľ�Ǯ���������������飬�����Ȼ����ӣ��˳�ʱҪ���۳����������ŷѣ�����Ҫ��һ�������������������
		<li>����<u>ħ��</u>��Ǯ�����ӣ����ǻ��һ����������������������˳�ʱ��֮
		<li>�������߳����ɵ��û����Զ�����<u>������</u> 
		<li>����һ���µ�����ǰ���������˳�ԭ�е����ɣ�<u>������</u>���� 
	</ui>	
	</td>
	<td valign=middle class=tablebody2 width="5%"></td>
</tr>
</table>
<br>
<table cellspacing=1 align=center class=tableborder1 width=<%=Forum_body(12)%>>
	<tr><td height=20 class=tablebody2 colspan=9 align=center><a href=z_menpaicompare.asp><B>�鿴�����ɶԱ�</b></a></td></tr>
  <tr>
    <th height=24>��������</th>
    <th height=24>����</th> 
    <th height=24>��������</th> 
    <th height=24>����</th> 
    <th height=24>������</th> 
    <th height=24>��Ǯ</th> 
    <th height=24>������</th> 
    <th height=24>�ܻ���</th> 
    <th height=24>����</th> 
  </tr>
  <%
'������ʾ  ��ʼ  
	dim rs1
	sql="select * from [GroupName]"
	set rs=conn.execute(sql)
	if rs.eof and rs.bof  then
		response.write "<tr><td height=21 class=tablebody1 colspan=8><font color=red>����̳��ʱû���κ�����</font></td></tr>"
	else 
		dim Zangmen,TotalArticle,TotalWealth,TotalUsernum,TotalScore
		dim menpai_setting

		do while not rs.eof
			TotalArticle=0
			TotalWealth=0
			TotalUsernum=0		
			TotalScore=0
			menpai_setting=split(rs("menpai_setting"),"|")
			Zangmen=trim(rs("Zangmen"))
			if Zangmen="" or isnull(Zangmen) then
				Zangmen="<font color=#FF6666>- ��ѡ�� -</font>"
			elseif Zangmen="-������·-" then
				Zangmen="<font color=#CC6666>������·</font>"
			else
				Zangmen="<a href=dispuser.asp?name="&htmlencode(Zangmen)&" title=""�鿴�������˵�����"" target=_blank>"&Zangmen&"</a>"
			end if	

			'ͳ��ָ�����ɵ�����������Ǯ�� �� ���������������ţ�
			set rs1=conn.execute("select sum(article),sum(userwealth),sum(userscore),count(*) from [user] where usergroup='"&rs("GroupName")&"'")
			if rs1(0)<>"" then TotalArticle=rs1(0)
			if rs1(1)<>"" then TotalWealth=rs1(1)
			if rs1(2)<>"" then TotalScore=rs1(2)
			if rs1(3)<>"" then TotalUsernum=rs1(3)
%>
<tr>
	<td valign=middle align=center height=21 class=tablebody1><a href="z_menpaiadd.asp?menu=1&groupname=<%=rs("groupname")%>" title="���ɼ��:<%=menpai_setting(7)%>"><%=rs("groupname")%></a></td>
	<td valign=middle align=center height=21 class=tablebody1><%=Zangmen%></td>
    <td valign=middle align=center height=21 class=tablebody1><%=rs("addtime")%></td>
    <td valign=middle align=center height=21 class=tablebody1><%=MenpaiAttri(menpai_setting(0))%></td>
    <td valign=middle align=center height=21 class=tablebody1><%=TotalUsernum%></td>
    <td valign=middle align=center height=21 class=tablebody1><%=TotalWealth%></td>
    <td valign=middle align=center height=21 class=tablebody1><%=TotalArticle%></td>
    <td valign=middle align=center height=21 class=tablebody1><%=TotalScore%></td>
    <td valign=middle align=center height=21 class=tablebody1><%if MenpaiState(rs("state"))<>"����" then%><%=MenpaiState(rs("state"))%><%else%><a href="z_menpaiadd.asp?menu=1&groupname=<%=rs("groupname")%>">�鿴</a><%end if%></td>
</tr>
<%			rs.movenext
		loop
		rs.close
	end if	  
'������ʾ  ����	

'���������ɲ���  ���뿪ʼ
	dim new_menpai_setting
	new_menpai_setting=split(conn.execute("select new_menpai_setting from [menpai]")(0),",")

	'set rs=conn.execute("select userwealth,usercp,userep,userpower,userclass from [user] where username='"&membername&"'")	
 	
	dim wumenwupai			'���ɲ�����������ɵ��ӵ�����
	set rs=conn.execute("select GroupName from [GroupName] where state=1")  
	if rs.eof and rs.bof then 
		wumenwupai="��������"
	else
		wumenwupai=rs(0)
	end if
	rs.close
	
	dim mymenpai,CanCreateMenpai
	CanCreateMenpai=true
	set rs=conn.execute("select GroupName from [GroupName] where Zangmen='"&membername&"'")
	if rs.eof and rs.bof then
		mymenpai=conn.execute("select UserGroup from [user] where UserID="&UserID)(0)
		if mymenpai=wumenwupai or mymenpai="��������" then
			mymenpai="<font color=navy>"&mymenpai&"</font>"
		else
			mymenpai="<font color=navy>"&mymenpai&"</font><font color=red>����</font>"
		end if	
	else
		mymenpai="<font color=navy>"&rs(0)&"</font><font color=red>����</font>"
	end if    		
%>
<tr><td height=21 class=tablebody2 colspan=9></td></tr>
<tr><th colspan=9 height=20>���뽨������</th></tr>
<tr>
	<td align=center height=21 class=tabletitle2>��Ŀ</td>
    <td align=center height=21 class=tabletitle2>��Ǯ</td>
	<td align=center height=21 class=tabletitle2>����</td>
	<td align=center height=21 class=tabletitle2>����</td>
	<td align=center height=21 class=tabletitle2>����</td>
	<td align=center height=21 class=tabletitle2>����</td>
	<td align=center height=21 class=tablebody2>�û��ȼ�</td>
	<td align=center height=21 class=tablebody2 colspan=2>����</td>
</tr>
<tr>
	<td align=center height=21 class=tablebody2>����Ҫ��</td>
    <td align=center height=21 class=tablebody1><%=new_menpai_setting(0)%>+<font color=red><%=new_menpai_setting(6)%></font></td>
	<td align=center height=21 class=tablebody1><%=new_menpai_setting(1)%></td>
	<td align=center height=21 class=tablebody1><%=new_menpai_setting(2)%></td>
	<td align=center height=21 class=tablebody1><%=new_menpai_setting(3)%></td>
	<td align=center height=21 class=tablebody1><%=new_menpai_setting(4)%></td>
	<td align=center height=21 class=tablebody1><%=GetUserTitle(new_menpai_setting(5))%>(<%=new_menpai_setting(5)%>��)</td>
	<td align=center height=21 class=tablebody1 colspan=2><%=wumenwupai%></td>
</tr>
<tr>
	<td align=center height=21 class=tablebody2>��������</td>
    <td align=center height=21 class=tablebody1><%=mymoney%></td>
	<td align=center height=21 class=tablebody1><%=myusercp%></td>
	<td align=center height=21 class=tablebody1><%=myuserep%></td>
	<td align=center height=21 class=tablebody1><%=mypower%></td>
	<td align=center height=21 class=tablebody1><%=myArticle%></td>
	<td align=center height=21 class=tablebody1><%=myClass%>(<%=GetUserClass(myClass)%>��)</td>
	<td align=center height=21 class=tablebody1 colspan=2><%=mymenpai%></td>
</tr>
<tr>
	<td align=center height=21 class=tablebody2>�Ƿ�����</td>
    <td align=center height=21 class=tablebody1><%if int(mymoney)>=(int(new_menpai_setting(0))+int(new_menpai_setting(6))) then%><font color=blue>��</font><%else%><%CanCreateMenpai=false%><font color=red>��</font><%end if%></td>
	<td align=center height=21 class=tablebody1><%if int(myusercp)>=int(new_menpai_setting(1)) then%><font color=blue>��</font><%else%><%CanCreateMenpai=false%><font color=red>��</font><%end if%></td>
	<td align=center height=21 class=tablebody1><%if int(myuserep)>=int(new_menpai_setting(2)) then%><font color=blue>��</font><%else%><%CanCreateMenpai=false%><font color=red>��</font><%end if%></td>
	<td align=center height=21 class=tablebody1><%if int(mypower)>=int(new_menpai_setting(3)) then%><font color=blue>��</font><%else%><%CanCreateMenpai=false%><font color=red>��</font><%end if%></td>
	<td align=center height=21 class=tablebody1><%if int(myArticle)>=int(new_menpai_setting(4)) then%><font color=blue>��</font><%else%><%CanCreateMenpai=false%><font color=red>��</font><%end if%></td>
	<td align=center height=21 class=tablebody1><%if int(GetUserClass(myClass))>=int(new_menpai_setting(5)) then%><font color=blue>��</font><%else%><%CanCreateMenpai=false%><font color=red>��</font><%end if%></td>
	<td align=center height=21 class=tablebody1 colspan=2><%if instr(mymenpai,wumenwupai)<>0 or instr(mymenpai,"��������")<>0 then%><font color=blue>��</font><%else%><%CanCreateMenpai=false%><font color=red>��</font><%end if%></td>
</tr>

<tr><td align=center height=21 class=tablebody1 colspan=9>
	<%if CanCreateMenpai then%>
		<a href="z_menpaiadd.asp?menu=8"><font color=blue>����������</font></a>	
	<%else%>
		<font color=red>��������һ�����������㴴��Ҫ��</font>	
	<%end if%>
	</td>
</tr>
<tr><td align=center height=21 class=tablebody2 colspan=9></td></tr>
</table>
<%
'���������ɲ���  ������� 
end sub
%>