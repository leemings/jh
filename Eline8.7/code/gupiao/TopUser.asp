<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html><head><title><%=Gupiao_Setting(5)%>-�ͻ�����</title>
<!--#include file="css.asp"-->
</head><body bgcolor="#ffffff" text="#000000" style="FONT-SIZE: 9pt" topmargin=5 leftmargin=0 oncontextmenu=self.event.returnValue=false>
	
<% 
	if request("action")="Reflash" then
		call Reflash()
	else	
		call main()
	end if	
response.Write "</body></html>"
CloseDatabase		'�ر����ݿ�  
'=====================================
sub main() 
	dim Title 
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
		Title="���ҽ��"
	else
		if request("paixu")="�ʲ�" then
			sql= "select * from �ͻ� order by �ʽ� desc"
			Title="�ʲ�����"	
		else 
			Title="���ʲ�����"
			sql= "select * from �ͻ� order by ���ʽ� desc"
		end if	
	end if					
%>
<table  width="98%" border=0 cellspacing=1 cellpadding=0 align=center bgcolor="#0066CC">
	<TR>
		<TD BACKGROUND="Images/topbg.gif" height=9 colspan=3>
	</TD>
	</TR>
	<tr>
		<td valign=center align=middle height=23 background="Images/Header.gif"><font size="2"><b>���д������а�</b></font></td>
	</tr>
	<tr><td bgcolor="#E4E8EF"> 
<br>
<table cellspacing=1 cellpadding=3 align=center width="97%" bgcolor="#0066CC">
	<tr>
		<td align=left valign=middle background="Images/title.gif" height="21"> &nbsp;<a href="stock.asp">��Ʊ���״���</a> | <a href="?paixu=���ʲ�">���ʽ�����</a> | <a href="?paixu=�ʲ�">�ʽ�����</a> | <a href=javascript:history.go(-1)>������һҳ</a></td>
	</tr> 
</table>
<br>

<table border="0" width="97%" bgcolor="#0066CC" cellspacing="1" cellpadding="3" align="center">
	<tr>
		<td valign=center align=middle height=23 background="Images/Header.gif" colspan="11"><b><%=Title%></b></td>
	</tr>
	<tr bgcolor="#ffffff"> 
	<%if request("action")="search" then%>
	<td colspan="11" height=18><font color=red>����������£�</font>����<a href="?paixu=<%=request("paixu")%>">[�������а�]</a></td>  
	<%else%>	
	<form name="form1" method="post" action="?action=search">
		<td colspan="11" height=18>���ٲ��ң�<input type="text"  name=username> ��<input type="checkbox"  name="usernamechk" value="yes" checked>�û�����ȫƥ�䣠<input type="submit" name="Submit" value="����"></td>
	</form>
	<%end if%>
	</tr>
	<tr bgcolor="#666666" align="center"> 
		<td background="Images/lan.gif"><font color=#ffffff>����</font></td>
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
			rs.PageSize = Gupiao_Setting(2) 
			rs.AbsolutePage=currentpage
			page_count=0
			totalrec=rs.recordcount
			do while (not rs.eof) and (not page_count = 20) 
%>
			<tr align=center bgcolor='#E4E6EF'>  
				<td><%if request("action")="search" then%>-<%else%><font color=red><%=Gupiao_Setting(2)*(currentpage-1)+page_count+1%></font><%end if%></td>
				<td><a href="dispu.asp?uid=<%=rs("id")%>&username=<%=rs("�ʺ�")%>"><%=rs("�ʺ�")%></a></td>
				<td align="right"><%=formatnumber(rs("�ʽ�"),0,true)%></td>
				<td align="right"><%=formatnumber(rs("���ʽ�"),0,true)%></td>
				<td><%=rs("�ֹ�����")%></td>
				<td><%=rs("��������")%></td>
				<td><%=rs("��������")%></td>
				<td><%=rs("������")%></td>
				<td><%=rs("������")%></td>
				<td><%=formatdatetime(rs("�������"),2)%></td>				
				<td><%if rs("����")=1 then%><font color=red>����<%else%>����<%end if%></td>
			</tr>
<%			
				page_count = page_count + 1		
				rs.movenext
			loop
			Pcount=rs.PageCount
%>
			<tr bgcolor="#666666" align="center"> 
				<td background="Images/lan.gif" colspan=2><font color=#ffffff>��λ</font></td>
				<td background="Images/lan.gif" align="right"><font color=#ffffff>Ԫ</font></td>
				<td background="Images/lan.gif" align="right"><font color=#ffffff>Ԫ</font></td>
				<td background="Images/lan.gif"><font color=#ffffff>��</font></td>
				<td background="Images/lan.gif"><font color=#ffffff>��</font></td> 
				<td background="Images/lan.gif"><font color=#ffffff>��</font></td> 
				<td background="Images/lan.gif"><font color=#ffffff>��</font></td>
				<td background="Images/lan.gif"><font color=#ffffff>��</font></td> 
				<td background="Images/lan.gif"><font color=#ffffff>��-��-��</font></td> 		
				<td background="Images/lan.gif"><font color=#ffffff>-</font></td> 
			</tr>
			<tr><td colspan=11 background="images/title.gif" align=right>
			<table border="0" cellpadding="0" cellspacing="0" width="100%"><tr>
			<td align=left>����<font color=blue><%=totalrec%></font>������ÿҳ<font color=blue><%=Gupiao_Setting(2)%></font>������<font color=red><%=currentpage%></font>ҳ/��<font color=blue><%=Pcount%></font>ҳ</td>
			<td align=right>��ҳ��
<%
			if currentpage > 4 then
				response.write "<a href=""?page=1&paixu="&request("paixu")&""">[1]</a> ..."
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
				response.write " <a href=""?page="&i&"&paixu="&request("paixu")&""">["&i&"]</a>"
				end if
			end if
			next
			if currentpage+3 < Pcount then 
				response.write "... <a href=""?page="&Pcount&"&paixu="&request("paixu")&""">["&Pcount&"]</a>"
			end if
			response.Write "</td></tr></table></td></tr>"		
		end if
		rs.close
		set rs=nothing	
%>
 </table>
 <br>
</td></tr>
<tr><td height=32 background="images/footer.gif" align=center><a href="stock.asp">���ع���</a></td></tr>
</table>
<% 
end sub

sub Reflash()
	if session("GP_ReflashFlag")<>"" then
		errmess="<li>���ڵ������Ѿ������µ��ˣ��벻Ҫ���ˢ��"
		call endinfo(1) 
	else
		dim rst,TotalFund,StockTypeNum
		set rs=conn.execute("select id from [�ͻ�] order by id")
		do while not rs.eof 
			TotalFund=0 :	StockTypeNum=0	
			sql="select d.�ֹ���,g.��ǰ�۸� from [��] d inner join [��Ʊ] g on d.sid=g.sid where d.uid="&rs(0)
			set rst=conn.execute(sql)
			do while not rst.eof
				TotalFund=TotalFund+rst(0)*rst(1)
				StockTypeNum=StockTypeNum+1
				rst.movenext
			loop
			conn.execute("update [�ͻ�] set ���ʽ�=�ʽ�+"&TotalFund&",�ֹ�����="&StockTypeNum&" where id="&rs(0))
			rs.movenext
		loop
		session("GP_ReflashFlag")="haveFlash"
		response.Redirect Request.ServerVariables("HTTP_REFERER")
	end if		
end sub
%>