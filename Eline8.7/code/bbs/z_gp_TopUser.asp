<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_gp_conn.asp"-->
<!--#include file="z_gp_Const.asp"-->
<%
stats="��������"
call nav()
call head_var(0,0,GuPiao_Setting(5),"z_gp_gupiao.asp")
response.write "<table class=tableborder1 cellspacing=1 cellpadding=3 align=center border=0 width=" & Forum_body(12) & ">"
if request("action")="Reflash" then
	call Reflash()
else	
	call main()
end if	
response.write "</table>"
call activeonline()
call footer()
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
			sql= "select * from [KeHu] where ZhangHao='"&username&"' and suoding<2"
		else	
			sql= "select * from [KeHu] where ZhangHao like '%"&username&"%' and suoding<2"         
		end if
		Title="���ҽ��"
	else
		if request("paixu")="�ʲ�" then
			sql= "select * from [KeHu] where suoding<2 order by ZiJin desc"
			Title="�ʲ�����"	
		else 
			Title="���ʲ�����"
			sql= "select * from [KeHu] where suoding<2 order by ZongZiJin desc"
		end if	
	end if%>
	<tr>
		<th valign=center align=middle height=25 colspan="11"><b><%=Title%></b></th>
	</tr>
	<%if request("action")="search" then%>
		<td class=tablebody2 colspan="11" height=18><font color=red>����������£�</font>����<a href="?paixu=<%=request("paixu")%>">[�������а�]</a></td>  
	<%else%>	
		<form name="form1" method="post" action="?action=search">
		<td class=tablebody2 colspan="11" height=18>���ٲ��ң�<input type="text"  name=username> ��<input type="checkbox"  name="usernamechk" value="yes" checked>�û�����ȫƥ�䣠<input type="submit" name="Submit" value="����"></td>
		</form>
	<%end if%>
	</tr>
	<tr align="center"> 
		<th>����</th>
		<th>�ʻ�</th>
		<th>�ʽ�</th>
		<th>���ʽ�</th>
		<th>�ֹ�����</th>
		<th>��������</th> 
		<th>��������</th> 
		<th>������</th>
		<th>������</th> 
		<th>�������</th> 		
		<th>״̬</th> 
	</tr>
	<%dim currentpage,page_count,Pcount
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
	rs.open sql,gp_conn,1,1
	if rs.eof and rs.bof then
		response.write "<tr><td class=tablebody1 colspan=11> û���ҵ��κι���</td></tr>"
	else
		rs.PageSize = Gupiao_Setting(2) 
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		do while (not rs.eof) and (not page_count = rs.PageSize)%>
			<tr align=center>  
				<td class=tablebody1><%if request("action")="search" then%>-<%else%><font color=red><%=Gupiao_Setting(2)*(currentpage-1)+page_count+1%></font><%end if%></td>
				<td class=tablebody1><a href="z_gp_Dispu.asp?uid=<%=rs("id")%>&username=<%=rs("ZhangHao")%>"><%=rs("ZhangHao")%></a></td>
				<td class=tablebody1 align="right"><%=formatnumber(rs("ZiJin"),0,true)%>&nbsp;</td>
				<td class=tablebody1 align="right"><%=formatnumber(rs("ZongZiJin"),0,true)%>&nbsp;</td>
				<td class=tablebody1><%=rs("ChiGuZhongLei")%></td>
				<td class=tablebody1 align="right"><%=rs("JinRiMaiRu")%>&nbsp;</td>
				<td class=tablebody1 align="right"><%=rs("JinRiMaiChu")%>&nbsp;</td>
				<td class=tablebody1 align="right"><%=rs("ZongMaiRu")%>&nbsp;</td>
				<td class=tablebody1 align="right"><%=rs("ZongMaiChu")%>&nbsp;</td>
				<td class=tablebody1><%=formatdatetime(rs("ZuiHouRiQi"),2)%></td>				
				<td class=tablebody1><%if rs("SuoDing")=1 then%><font color=red>����<%else%>����<%end if%></td>
			</tr>
			<%page_count = page_count + 1		
			rs.movenext
		loop
		Pcount=rs.PageCount%>
		<tr>
			<td class=tablebody2 colspan=3 align=left>���� <font color=blue><%=totalrec%></font> ������ÿҳ <font color=blue><%=rs.PageSize%></font> ������ <font color=red><%=currentpage%></font> ҳ/�� <font color=blue><%=Pcount%></font> ҳ</td>
			<td class=tablebody2 colspan=8 align=right>��ҳ��<%call disppagenum(currentpage,pcount,"""?page=","&paixu="&request("paixu")&"""")%>&nbsp;&nbsp;&nbsp;&nbsp;</td>
		</tr>
	<%end if
	rs.close
	set rs=nothing%>
	<tr>
		<th colspan="11" height=25 align=center><a href="z_gp_Gupiao.asp">���ع���</a></th>
	</tr>
<%end sub

sub Reflash()
	if session("GP_ReflashFlag")<>"" then
		errmess="<li>���ڵ������Ѿ������µ��ˣ��벻Ҫ���ˢ��"
		call endinfo(1) 
	else
		dim rst,TotalFund,StockTypeNum
		set rs=gp_conn.execute("select id from [KeHu] where suoding<2 order by id")
		do while not rs.eof 
			TotalFund=0 :	StockTypeNum=0	
			sql="select d.ChiGuShu,g.DangQianJiaGe from [DaHu] d inner join [GuPiao] g on d.sid=g.sid where d.uid="&rs(0)
			set rst=gp_conn.execute(sql)
			do while not rst.eof
				TotalFund=TotalFund+rst(0)*rst(1)
				StockTypeNum=StockTypeNum+1
				rst.movenext
			loop
			gp_conn.execute("update [KeHu] set ZongZiJin=ZiJin+"&TotalFund&",ChiGuZhongLei="&StockTypeNum&" where id="&rs(0))
			rs.movenext
		loop
		session("GP_ReflashFlag")="haveFlash"
		response.Redirect Request.ServerVariables("HTTP_REFERER")
	end if		
end sub
%>