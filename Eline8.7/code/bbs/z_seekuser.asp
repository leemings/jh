<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_plus_check.asp"-->
<% 
dim nam
stats="��Ա��¼���"
call nav() 
call head_var(2,0,"","")
if not founduser or membername="" then
	errmsg=errmsg+"<br>"+"<li>����û��<a href=login.asp>��¼��̳</a>������ʹ��������ܡ����������û��<a href=reg.asp>ע��</a>������<a href=reg.asp>ע��</a>��"
	founderr=true
	call dvbbs_error()
else%>
<table class=tableborder1 cellspacing=1 cellpadding=3 align=center>
	<tr> 
		<th height="25" align="center" colspan=7>��Ա��ѯ</th>
	</tr>
	<form method="POST" action="?">
	<tr> 
		<td class=tablebody1 valign=middle height=50 align=center>������Ҫ��ѯ�Ļ�Ա: <input type="text" name="nam1" size="12">  <input type="submit" value="�ύ" name="B1"></td>
	</tr>
	</form>
</table>
<br>
<%nam=checkstr(trim(request("nam1")))
if nam<>"" then
	call hy(nam)
end if
end if
call activeonline()
call footer()

sub hy(nam)
	dim page,pagecount,totalrec,pagesize
	dim GlTime,sex
	
	pagesize=20
	page=trim(request("page"))
	if page="" then
		page=1
	else
		page=cint(page)
	end if
	set rs = server.CreateObject ("adodb.recordset")
	sql="select username,adddate,Sex,Userclass,UserGroup,lastlogin from [user] where username like '%"&nam&"%' order by datediff('d',lastlogin,date()) "
	rs.open sql,conn,1,1
	totalrec=rs.recordcount
	if totalrec mod pagesize=0 then
		pagecount=totalrec \ pagesize
	else
		pagecount=totalrec \ pagesize+1
	end if
	if page > pagecount then page = pagecount
	if page < 1 then page=1
	rs.PageSize = pagesize%>
	<table Class=TableBorder1 cellspacing="1" cellpadding="1" align="center" >
		<tr>
			<th height="25" align="center" colspan=7>��Ա��¼���</td>
		</tr>
  	<%if rs.EOF or rs.BOF then%>
  		<tr><td align="center" height="25" class=tablebody1 colspan=7>û����˵�������Ļ�ԱŶ����</td></tr>
  	<%else%>
   		<tr height=23 align=center> 
    		<td class=tablebody2>����Ա������</td> 
    		<td class=tablebody2>��ע�����ڡ�</td>
    		<td class=tablebody2>����ʧ������</td>
				<td class=tablebody2>����Ա�Ա�</td>
    		<td class=tablebody2>����Ա�ȼ���</td>
				<td class=tablebody2>����Ա���ɡ�</td>
    		<td class=tablebody2>��������Ϣ��</td>
			</tr>
  		<%rs.AbsolutePage=page
  		i=0
  		do while not rs.EOF%>
  			<tr height=23 align=center> 
  				<td class=tablebody1> <a href="dispuser.asp?name=<%=rs(0)%>" target="_blank" title="�鿴<%=rs(0)%>�ĸ�������"><%=rs(0)%></a> </td>
  				<td class=tablebody1><%=rs(1)%></td>
					<% GlTime= datediff("d",rs(5),date()) %>
  				<td width="10%" class=tablebody1><b><font color=red><%=GlTime%></font></b> ��</td>
					<% if rs(2)=1 then
   					sex="<font color=#0066CC>��</font>"
   				elseif rs(2)=0 then
   					sex="<font color=red>Ů</font>" 
					end if %>
  				<td class=tablebody1><%=sex%></td>
  				<td class=tablebody1><%=rs(3)%></td>
  				<td class=tablebody1><%=rs(4)%></td>	
  				<td class=tablebody1><a target="_blank" href="messanger.asp?action=new&touser=<%=rs("username")%>"><img src="pic/message.gif" alt=�������� border="0"></a></td>
				</tr>
				<%	i=i+1
				rs.movenext
				if i>= rs.PageSize then exit do
			loop  
		end if %>
		<tr align="center"> 
			<td height="25" align=center class=tablebody2><a href="z_viewmanage.asp">[��������]</a></td>
			<td height="25" align=left class=tablebody2>&nbsp;�� <b><%=rs.RecordCount%></b> �� �� <b><%=page%></b> / <b><%=rs.PageCount%></b> ҳ</td>
			<td height="25" align=right class=tablebody2 colspan=5><%call disppagenum(page,pagecount,"?page=","&nam1="&nam)%>&nbsp;</td>
		</tr>
	</table>
	<%rs.Close
	set rs = nothing 
end sub
%>
