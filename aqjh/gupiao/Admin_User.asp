<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html><head><title><%=Gupiao_Setting(5)%>-���п������</title>
<!--#include file="css.asp"-->
</head><body bgcolor="#ffffff" text="#000000" style="FONT-SIZE: 9pt" topmargin=5 leftmargin=0 oncontextmenu=self.event.returnValue=false>
<table  width="98%" border=0 cellspacing=1 cellpadding=0 align=center bgcolor="#0066CC">
	<TR>
		<TD BACKGROUND="Images/topbg.gif" height=9 colspan=3>
	</TD>
	</TR>
	<tr>
		<td valign=center align=middle height=23 background="Images/Header.gif"><font size="2"><b>���й�������</b></font></td>
	</tr>
	<tr><td bgcolor="#E4E8EF"> 
<% 
call AdminHead()
if not master then
	errmess="<li>��ҳΪ����Աר�ã���û�й���ҳ��Ȩ�ޣ�"
	call endinfo(1)
else 
	select case request("action")
		case "EditUser"
			call EditUser()
		case "SaveEdit"		
			call SaveEdit()
		case "del"
			call DelUser()	
		case else
			call main()
	end select
end if
%>
<br>
</td></tr>
<tr><td height=32 background="images/footer.gif" valign=middle></td></tr>
</table>
</body></html>
<%
CloseDatabase		'�ر����ݿ�
'=====================================
sub main()
	if request("action")="search" then
		dim username
		username=trim(replace(request("username"),"'",""))
		if username="" then
			errmess="<li>��������ҹؼ���"
			rurl="?"
			call endinfo(2)
			exit sub
		end if	
		if request("usernamechk")="yes" then
			sql= "select * from �ͻ� where �ʺ�='"&username&"'"
		else	
			sql= "select * from �ͻ� where �ʺ� like '%"&username&"%'"         
		end if
	else 
		sql= "select * from �ͻ� order by �ʽ� desc"	
	end if					
%>
<table border="0" width="97%" bgcolor="#0066CC" cellspacing="1" cellpadding="3" align="center">
	<tr>
		<td valign=center align=middle height=23 background="Images/Header.gif" colspan="11"><b>�ͻ�����</b></td>
	</tr>
	<tr bgcolor="#ffffff"> 
	<%if request("action")="search" then%>
	<td colspan="11" height=18><font color=red>����������£�</font>����<a href="?">[���ؿͻ��б�]</a></td>
	<%else%>	
	<form name="form1" method="post" action="?action=search">
		<td colspan="11" height=18>���ٲ��ң�<input type="text"  name=username> ��<input type="checkbox"  name="usernamechk" value="yes" checked>�û�����ȫƥ�䣠<input type="submit" name="Submit" value="��ѯ�ͻ�"></td>
	</form>
	<%end if%>
	</tr>
	<tr bgcolor="#666666" align="center"> 
		<td background="Images/lan.gif"><font color=#ffffff>�ʻ�</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>�ʽ�</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>���ʽ�</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>�ֹ�����</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>��������</font></td> 
		<td background="Images/lan.gif"><font color=#ffffff>��������</font></td> 
		<td background="Images/lan.gif"><font color=#ffffff>������</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>������</font></td> 
		<td background="Images/lan.gif"><font color=#ffffff>�������</font></td> 		
		<td background="Images/lan.gif"><font color=#ffffff>״̬</font></td> 
		<td background="Images/lan.gif"><font color=#ffffff>����</font></td>
	</tr>
<%
		dim currentpage,page_count,Pcount
		dim totalrec,endpage
		currentPage=request("page")
		if currentpage="" or not isnumeric(currentpage) then
			currentpage=1
		else
			currentpage=clng(currentpage)
			if err then
				currentpage=1
				err.clear
			end if
		end if
		Set rs= Server.CreateObject("ADODB.Recordset")
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write "<tr><td colspan=11 bgcolor=#E4E8EF> û���ҵ��κι���</td></tr>"
		else
			rs.PageSize = 20 
			rs.AbsolutePage=currentpage
			page_count=0
			totalrec=rs.recordcount
			do while (not rs.eof) and (not page_count = 20)
%>
			<tr align=center> 
				<td bgcolor="#FFFFFF"><a href="?action=EditUser&username=<%=rs("�ʺ�")%>"><%=rs("�ʺ�")%></a></td>
				<td bgcolor="#FFFFFF" align="right"><%=formatnumber(rs("�ʽ�"),2,true)%></td>
				<td bgcolor="#FFFFFF" align="right"><%=formatnumber(rs("���ʽ�"),2,true)%></td>
				<td bgcolor="#FFFFFF"><%=rs("�ֹ�����")%></td>
				<td bgcolor="#FFFFFF"><%=rs("��������")%></td>
				<td bgcolor="#FFFFFF"><%=rs("��������")%></td>
				<td bgcolor="#FFFFFF"><%=rs("������")%></td>
				<td bgcolor="#FFFFFF"><%=rs("������")%></td>
				<td bgcolor="#FFFFFF"><%=formatdatetime(rs("�������"),2)%></td>				
				<td bgcolor="#FFFFFF"><%if rs("����")=1 then%><font color=red>����<%else%>����<%end if%></td>
				<td bgcolor="#FFFFFF"><a href="?action=EditUser&username=<%=rs("�ʺ�")%>">�༭</a> | <a href="?action=del&username=<%=rs("�ʺ�")%>" onclick="javascript:{if(confirm('��ȷ��ɾ�� <%=rs("�ʺ�")%> ����ͻ���?')){return true;}return false;}">ɾ��</a></td>
			</tr>
<%			
				page_count = page_count + 1		
				rs.movenext
			loop
			Pcount=rs.PageCount
%>
			<tr bgcolor="#666666" align="center"> 
				<td background="Images/lan.gif"><font color=#ffffff>��λ</font></td>
				<td background="Images/lan.gif" align="right"><font color=#ffffff>Ԫ</font></td>
				<td background="Images/lan.gif" align="right"><font color=#ffffff>Ԫ</font></td>
				<td background="Images/lan.gif"><font color=#ffffff>��</font></td>
				<td background="Images/lan.gif"><font color=#ffffff>��</font></td> 
				<td background="Images/lan.gif"><font color=#ffffff>��</font></td> 
				<td background="Images/lan.gif"><font color=#ffffff>��</font></td>
				<td background="Images/lan.gif"><font color=#ffffff>��</font></td> 
				<td background="Images/lan.gif"><font color=#ffffff>��-��-��</font></td> 		
				<td background="Images/lan.gif"><font color=#ffffff>-</font></td> 
				<td background="Images/lan.gif"><font color=#ffffff>-</font></td>
			</tr>
			<tr><td colspan=11 background="images/title.gif" align=right>
			<table border="0" cellpadding="0" cellspacing="0" width="100%"><tr>
			<td align=left>����<font color=blue><%=totalrec%></font>������ÿҳ<font color=blue><%=rs.PageSize%></font>������<font color=red><%=currentpage%></font>ҳ/��<font color=blue><%=Pcount%></font>ҳ</td>
			<td align=right>��ҳ��
<%
			if currentpage > 4 then
				response.write "<a href=""?page=1"">[1]</a> ..."
			end if
			if Pcount>currentpage+3 then
			endpage=currentpage+3
			else
			endpage=Pcount
			end if
			for i=currentpage-3 to endpage
			if not i<1 then
				if i = clng(currentpage) then
				response.write " <font color=red>["&i&"]</font>"
				else
				response.write " <a href=""?page="&i&""">["&i&"]</a>"
				end if
			end if
			next
			if currentpage+3 < Pcount then 
				response.write "... <a href=""?page="&Pcount&""">["&Pcount&"]</a>"
			end if
			response.Write "</td></tr></table></td></tr>"		
		end if
		rs.close
		set rs=nothing	
%>
</table>
<% 
end sub
'---------------�༭�û�---------
sub EditUser()
	dim username
	username=trim(replace(request("username"),"'",""))
	if username="" then
		errmess="<li>����Ĳ�������ָ���û���"
		call endinfo(2)
		exit sub
	end if	
	sql= "select * from �ͻ� where �ʺ�='"&username&"'"
	set rs=conn.execute(sql)         
	if rs.eof then
		errmess="<li>���û��Ĳ�����"
		call endinfo(2)
	else		
%>
<table  width="420" height=300 border=0 cellspacing=1 cellpadding=3 align=center bgcolor="#0066CC">
	<TR>
	<TD BACKGROUND="images/topbg.gif" height=9 colspan="2"></TD>
	</TR>
	<tr>
		<td colspan="2" valign=center align=middle height=23 class=big background="images/header.gif">�༭ ����<font color=red><%=rs("�ʺ�")%></font>��������</td>
	</tr>
	<form method=post  action="?action=SaveEdit">
	<tr>
		<td BGCOLOR="#E4E4EF" width="40%">UID</td>
		<td bgcolor="#FFFFFF" width="60%">&nbsp;<%=rs("id")%><input type=hidden name=Uid value="<%=rs("id")%>"></td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF">�ʻ�</td>
		<td bgcolor="#FFFFFF">&nbsp;<font color=blue><%=rs("�ʺ�")%></font><input type=hidden name=username value="<%=rs("�ʺ�")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF">��������</td>
		<td bgcolor="#FFFFFF">&nbsp;<%=rs("��������")%></td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF">�������</td>
		<td bgcolor="#FFFFFF">&nbsp;<%=rs("�������")%></td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF">״̬</td>
		<td bgcolor="#FFFFFF"><input type=radio name=Locked value="0" <%if rs("����")<>1 then%> checked <%end if%>> ������<input type=radio name=Locked value="1" <%if rs("����")=1 then%> checked <%end if%>> ���� </td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF">�ʽ�</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=Crash value="<%=rs("�ʽ�")%>"></td>
	</tr>			
	<tr>
		<td BGCOLOR="#E4E4EF">���ʽ�</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=TotalFund value="<%=rs("���ʽ�")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF"><font color="#333333">�ֹ�����</font></td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=StockNum value="<%=rs("�ֹ�����")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF"><font color=gray>��������</font></td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=TodayBuy readonly value="<%=rs("��������")%>"></td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF"><font color=gray>��������</font></td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=TodaySale readonly value="<%=rs("��������")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF"><font color=gray>������</font></td> 
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=TotalBuy readonly value="<%=rs("������")%>"></td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF"><font color=gray>������</font></td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=TotalSale readonly value="<%=rs("������")%>"></td>
	</tr>
	<tr><td colspan="2" bgcolor="#FFFFFF" align="center"><input type=submit name=submit value=ִ���޸�></td>
	</form>
</table>	
<%		
	end if
end sub
'---------------�����޸�---------
sub SaveEdit()
	dim username,Uid
	username=trim(replace(request.Form("username"),"'",""))
	Uid=request.Form("Uid")
	if username="" or (not isnumeric(uid)) then
		errmess="<li>����Ĳ������������еĲ����Ƿ�"
		call endinfo(2)
		exit sub
	end if	
	sql= "select * from �ͻ� where id="&Uid&""         
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs("�ʽ�")=request.form("Crash")
	rs("���ʽ�")=request.form("TotalFund")
	rs("����")=request.form("Locked")
	rs("�ֹ�����")=request.form("StockNum")
	rs("��������")=request.form("TodayBuy")
	rs("��������")=request.form("TodaySale")
	rs("������")=request.form("TotalBuy")
	rs("������")=request.form("TotalSale")		
	rs.update
	rs.close
	sucmess="<li>�û� <font color=blue>"&username&"</font> �����ϸ��³ɹ�"
	rUrl="?"
	call endinfo(2)
end sub

sub DelUser()
		
end sub
%>