<%
function iif(expression,returntrue,returnfalse)
	if expression=0 then
		iif=returnfalse
	else
		iif=returntrue
	end if
end function

'ȡ������������Ǹ�
function Max(expression1,expression2)
	if expression1+0>=expression2+0 then
		Max=expression1
	else
		Max=expression2
	end if
end function

Rem ����HTML����
function HTMLEncode(fString)
	if not isnull(fString) then
		fString = replace(fString, ">", "&gt;")
		fString = replace(fString, "<", "&lt;")
	
		fString = Replace(fString, CHR(32), "&nbsp;")
		fString = Replace(fString, CHR(9), "&nbsp;")
		fString = Replace(fString, CHR(34), "&quot;")
		fString = Replace(fString, CHR(39), "&#39;")
		fString = Replace(fString, CHR(13), "")
		fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
		fString = Replace(fString, CHR(10), "<BR> ")
	
		HTMLEncode = fString
	end if
end function
'-------------------------------��Ϣ��ʾ-------------------------------
'flag=0 ��ʾ�ر�ҳ��	flag=1 ��ʾ���ع�Ʊ���״���		flag=2	��ʾ������һҳ
sub endinfo(flag) 
%>
<br>
<table  width="75%" border=0 cellspacing=1 cellpadding=3 align=center bgcolor="#0066CC">
	<TR><TD BACKGROUND="Images/topbg.gif" height=9></TD></TR> 
	<tr><td valign=center align=middle height=23 background="Images/Header.gif"><b>��Ʊ������Ϣ��ʾ</b></td>
	<tr><td height=1 bgcolor="#E4E8EF">
<%
	if errmess<>"" then
		response.Write "<b>��������Ŀ���ԭ��</b><br><br><li>���Ƿ���ϸ�Ķ��˹���ָ�ϣ���������û�е�½���߲�����ʹ�õ�ǰ���ܵ�Ȩ��"	
		response.Write "<font color=red>"&errmess&"</font>"
	else
		response.Write "<b>�����ɹ���</b><br><br><li>��ӭ���ٰ��齭����Ʊ�������ģ��뷵�ؽ�����������"
		response.Write "<font color=navy>"&sucmess&"</font>"
	end if		
%>	
	<br></td></tr> 
	<TR><TD BACKGROUND="Images/title.gif" height=21 align="center">
<%
	if flag=0 then
		response.Write "<a href=""javascript:window.close()"">[�뿪��Ʊ��������]</a>����<a href=""#"" onclick=""window.close()"">[�رձ�ҳ]</a>"
	else
		if rUrl="" or isnull(rurl) then
			rUrl=Request.ServerVariables("HTTP_REFERER")
		end if
		response.Write "<a href=""stock.asp"" class=cblue>[���ع��д���]</a>&nbsp;&nbsp;"	
		response.Write "<a href="""&rUrl&""" class=cblue>[������һҳ]</a>"
	end if
%></TD></TR> 
</table>
<%
end sub

sub AdminHead()
%>
<br>
<table cellspacing=1 cellpadding=3 align=center width="97%" bgcolor="#0066CC">
	<tr>
		<td align=left valign=middle background="Images/title.gif" height="21"> &nbsp;<a href="stock.asp">��Ʊ���״���</a> �� <a href="Admin_Setting.asp">��������</a> �� <a href="Admin_Gupiao.asp">��Ʊ����</a> �� <a href="Admin_User.asp">�ͻ�����</a> �� <a href="Announcements.asp">�������</a> �� <a href="Admin_Data.asp">���ݿ����</a> �� <a href=javascript:history.go(-1)>������һҳ</a></td>
	</tr> 
</table>
<br>
<%
end sub
'----------------------------------------
Rem ���߹���
sub GPOnline()
	dim statuserid
	statuserid=replace(Request.ServerVariables("REMOTE_HOST"),".","")
	if membername="" or MyUserID=0 then
		Dim TempUsername
		TempUsername=iif(membername="","����",membername)
		session("GP_UserID")=statuserid
		sql="select id from online where id="&cstr(session("GP_UserID"))
		set rs=conn.execute(sql)
		if rs.eof and rs.bof then
			sql="insert into online(id,username,Uid,ip,startime,lastimebk,stats,userhidden) values ("&statuserid&",'"&TempUsername&"',0,'"&Request.ServerVariables("REMOTE_HOST")&"',Now(),Now(),'"&replace(stats,"'","")&"',2)"
		else
			sql="update online set lastimebk=Now(),stats='"&replace(stats,"'","")&"' where id="&cstr(session("GP_UserID"))
		end if
		conn.execute(sql)
	else
		sql="select id from online where Uid="&MyUserID
		set rs=conn.execute(sql)
		if rs.eof and rs.bof then
			sql="insert into online(id,username,Uid,ip,startime,lastimebk,stats,userhidden) values ("&statuserid&",'"&membername&"',"&MyUserID&",'"&Request.ServerVariables("REMOTE_HOST")&"',Now(),Now(),'"&replace(stats,"'","")&"',"&userhidden&")"		
		else
			sql="update online set lastimebk=Now(),stats='"&replace(stats,"'","")&"' where Uid="&MyUserID
		end if
		conn.execute(sql)
		conn.execute("update [�ͻ�] set �������=now() where id="&MyUserID)
		rs.close
		if session("GP_UserID")<>"" then
			Conn.Execute("delete from online where id="&session("GP_UserID"))
			session("GP_UserID")=""
		end if
	end if
	set rs=nothing
	Rem ɾ����ʱ�û�
	sql="Delete FROM online WHERE DATEDIFF('s', lastimebk, now()) > "&Gupiao_Setting(9)&"*60"
	Conn.Execute sql
end sub
'��������
function online(boardid)
	online=conn.execute("Select count(*) from online")(0)
	if isnull(online) then online=0
end function 
'��ʾ���߹����б�
sub onlineuser(online_u,online_g)
response.write "<tr><td colspan=2><table cellpadding=0 cellspacing=0 border=0 width=""100%"" style=""word-break:break-all;""><tr height=20 valign=middle>"
dim userip,NowStats
if cint(online_u)=1 then		'������ʾ��Ա��Ϣ
	i=0
	'�û���Ϣ
	sql="select username,ip,stats,userhidden,Uid,startime,lastimebk from online where uid>=0"
	set rs=conn.execute(sql)
	do while not rs.eof
	NowStats="Ŀǰλ�ã�" & server.HTMLEncode(rs(2)) &"<br>"
	if master then
		userip="��ʵ�ɣУ�" & rs(1)
	else
		userip="��ʵ�ɣУ������ñ���"
	end if
	
	if membername=rs(0) then		'�Լ�
		response.write "<td background=""images/title.gif"" width=""14%"">&nbsp;<a href=dispu.asp?uid="&rs(4)&" target=_blank class=cblue title="""& NowStats & UserIP &""">"&server.htmlencode(rs(0))&"</font></a></td>"
	else
		if rs(3)=1 then			'������������
			if master then
				response.write "<td background=""images/title.gif"" width=""14%"">&nbsp;<a href=dispu.asp?uid="&rs(4)&" target=_blank title="""& NowStats & UserIP & "<br>��ǰ״̬������"">"&server.htmlencode(rs(0))&"</a></td>"
			else
				response.write "<td background=""images/title.gif"" width=""14%"">&nbsp;<a href=# class=cgray target=_blank title="""& NowStats & UserIP &""">�������</a></td>"
			end if
		else
			response.write "<td background=""images/title.gif"" width=""14%"">&nbsp;<a href=dispu.asp?uid="&rs(4)&" target=_blank title="""& NowStats & UserIP &""">"&server.htmlencode(rs(0))&"</a></td>"
		end if
	end if
	if i=6 then response.write "</tr><tr height=20 valign=middle>"
	if i>6 then 
		i=1
	else
		i=i+1
	end if
	rs.movenext
	loop
end if
response.write "</tr></TABLE>"
set rs=nothing
end sub
'-------------------------------��Ʊ������-------------------------------
sub CloseGuPiao(expression)
	if expression=1 then
%>
	<title><%=Gupiao_Setting(5)%>-�����Ѿ�������</title>
	<style type=text/css><!--
		td {  font-family: ����; font-size: 9pt}-->
	</style>
	<br>
	<table cellspacing=1 cellpadding=3 align=center width="75%" bgcolor="#0066CC" border=0>
	<TR><TD BACKGROUND="Images/topbg.gif" height=9 colspan=3></TD></TR> 
	<tr><td height=23 background="Images/header.gif" align="center">��Ʊ����֪ͨ</td></tr>
	<tr><td width="100%" bgcolor="#FFFFFF" height="50" valign="middle">
						�װ��Ĺ��� <font color=blue><%=membername%></font>�����ã�<br>
						&nbsp;&nbsp;&nbsp;&nbsp;���ڹ�Ʊ�г��Ѿ������ˣ�����ÿ��� <%=split(Gupiao_Setting(4),"||")(0)%>:00��<%=split(Gupiao_Setting(4),"||")(1)%>:00 ������лл������ 
	</td></tr>
	<tr><td align=center height=26 bgcolor="#E4E8EF"><a href="javascript:window.close()">[�뿪��Ʊ��������]</a>����<a href="#" onClick="window.close()">[�رձ�ҳ]</a></td></tr> 
	</table>
	<%elseif expression=0 then%>
	<title><%=Gupiao_Setting(5)%>-������ʱ�ر�</title>
	<style type=text/css><!--
		td {  font-family: ����; font-size: 9pt}--> 
	</style>
	<br><br>
	<table cellspacing=1 cellpadding=3 align=center width="75%" bgcolor="#0066CC" border=0>
		<TR><TD BACKGROUND="Images/topbg.gif" height=9 colspan=3></TD></TR> 
		<tr><td height=23 background="Images/header.gif" align="center">��ӭ���� <%=Gupiao_Setting(5)%></td></tr>
		<tr>
			<td bgcolor="#FFFFFF"><br><font color="#FF3366"> �� ���ڹ��������ڹر�״̬�У�ԭ�����£�</font><br><br>&nbsp;&nbsp;&nbsp;<%=StopReadme%><br><br></td>
		</tr>
		<tr><td align=center height=26 bgcolor="#E4E8EF"><a href="javascript:window.close()">[�뿪���н�������]</a>����<a href="#" onClick="window.close()">[�رձ�ҳ]</a></td></tr> 
	</table>
	<%else%>
	<title><%=Gupiao_Setting(5)%>-��ˢ�¹��ܿ���</title>
	<style type=text/css><!--
		td {  font-family: ����; font-size: 9pt}--> 
	</style> 
	<META http-equiv=Content-Type content=text/html; charset=gb2312> 
	<br><br>
	<table cellspacing=1 cellpadding=3 align=center width="75%" bgcolor="#0066CC" border=0>
		<TR><TD BACKGROUND="Images/topbg.gif" height=9 colspan=3></TD></TR> 
		<tr><td height=23 background="Images/header.gif" align="center">��ˢ�»���</td></tr>
		<tr>
			<td bgcolor="#FFFFFF">
				<br>&nbsp;&nbsp;<font color="#FF3366">��ҳ�������˷�ˢ�»��ƣ��벻Ҫ�� <font color=blue><%=Gupiao_Setting(7)%></font> ��������ˢ�±�ҳ��</font><br><br> 
				&nbsp;&nbsp;<font color=blue><%=Gupiao_Setting(7)%></font> ��֮�󽫻��Զ���ҳ�棬���Ժ󡭡�<br><br>
			</td>  
		</tr> 
		<tr><td align=center height=26 bgcolor="#E4E8EF" id="TdReFlash"></td></tr> 
	</table>
	<meta HTTP-EQUIV=REFRESH CONTENT="<%=Gupiao_Setting(7)%>">
	<script language="VBScript"> 
	<!--
		TimeLimit=<%=Gupiao_Setting(7)%>
		call GetSec()
		function GetSec()
			TimeLimit=TimeLimit-1
			TdReFlash.innerHTML = "<font color=blue>"&TimeLimit&"</font>"
			if TimeLimit<= 0 then
				TdReFlash.innerHTML="<font color=red>���ڴ�ҳ�桭��</font>"
			else
				setTimeout "GetSec()",1000 
			end if
		end function
	//-->	
	</script>		
<%
	end if
end sub 

sub History()
	dim rst,TongJi
	set rst=conn.execute("select TongJi,��ǰ�۸�,sid,ʣ��ɷ�,IniTradeNum from [��Ʊ] order by sid")
		
	do while not rst.eof
		if rst(3)>rst(4) then
			sql="������=IniTradeNum "
		elseif rst(3)>0 then
			sql="������=ʣ��ɷ� "
		else
			sql="������=0,״̬='��' "
		end if	
		TongJi=rst(0)
		if TongJi="" or isnull(TongJi) then TongJi="0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"
		TongJi=mid(TongJi,instr(TongJi,"|")+1)&"|"&formatnumber(rst(1),2,-1)
		conn.execute("update [��Ʊ] set TongJi='"&TongJi&"',���̼۸�=��ǰ�۸�,�ɽ���=0,�������=0,��������=0,TodayWave=0,"&sql&" where sid="&rst(2))
		rst.movenext
	loop
end sub


%>