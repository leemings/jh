<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_gp_conn.asp"-->
<!--#include file="z_gp_Const.asp"-->
<%
stats="�ͻ�����"
call nav()
call head_var(0,0,GuPiao_Setting(5),"z_gp_gupiao.asp")
if not master then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
else
	call AdminHead()%>
	<table class=tableborder1 cellspacing=1 cellpadding=3 align=center border=0 width="<%=Forum_body(12)%>">
	<%select case request("action")
		case "EditUser"
			call EditUser()
		case "SaveEdit"		
			call SaveEdit()
		case "del"
			call DelUser()
		case else
			call main()
	end select%>
	</table>
<%end if
call footer()
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
			sql= "select * from [KeHu] where ZhangHao='"&username&"' and suoding<2"
		else	
			sql= "select * from [KeHu] where ZhangHao like '%"&username&"%' and suoding<2"         
		end if
	else 
		sql= "select * from [KeHu] where suoding<2 order by ZongZiJin desc"	
	end if%>
	<tr>
		<th height=25 colspan="11"><b>�ͻ�����</b></td>
	</tr>
	<tr> 
		<%if request("action")="search" then%>
			<td class=tablebody2 colspan="11" height=18><font color=red>����������£�</font>����<a href="?">[���ؿͻ��б�]</a></td>
		<%else%>	
			<form name="form1" method="post" action="?action=search">
			<td class=tablebody2 colspan="11" height=18>���ٲ��ң�<input type="text"  name=username> ��<input type="checkbox"  name="usernamechk" value="yes" checked>�û�����ȫƥ�䣠<input type="submit" name="Submit" value="��ѯ�ͻ�"></td>
			</form>
		<%end if%>
	</tr>
	<tr align="center" height=25> 
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
		<th>����</th>
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
				<td class=tablebody1><a href="?action=EditUser&username=<%=rs("ZhangHao")%>"><%=rs("ZhangHao")%></a></td>
				<td class=tablebody1 align="right"><%=formatnumber(rs("ZiJin"),2,true)%>&nbsp;</td>
				<td class=tablebody1 align="right"><%=formatnumber(rs("ZongZiJin"),2,true)%>&nbsp;</td>
				<td class=tablebody1 align="right"><%=rs("ChiGuZhongLei")%>&nbsp;</td>
				<td class=tablebody1 align="right"><%=rs("JinRiMaiRu")%>&nbsp;</td>
				<td class=tablebody1 align="right"><%=rs("JinRiMaiChu")%>&nbsp;</td>
				<td class=tablebody1 align="right"><%=rs("ZongMaiRu")%>&nbsp;</td>
				<td class=tablebody1 align="right"><%=rs("ZongMaiChu")%>&nbsp;</td>
				<td class=tablebody1><%=formatdatetime(rs("ZuiHouRiQi"),2)%></td>				
				<td class=tablebody1><%if rs("SuoDing")=1 then%><font color=red>����<%else%>����<%end if%></td>
				<td class=tablebody1><a href="?action=EditUser&username=<%=rs("ZhangHao")%>">�༭</a> | <a href="?action=del&username=<%=rs("ZhangHao")%>" onclick="javascript:{if(confirm('��ȷ��ɾ�� <%=rs("ZhangHao")%> ����ͻ���?')){return true;}return false;}">ɾ��</a></td>
			</tr>
			<%page_count = page_count + 1		
			rs.movenext
		loop
		Pcount=rs.PageCount%>
		<tr>
			<td class=tablebody2 colspan=3 align=left> ���� <font color=blue><%=totalrec%></font> ������ÿҳ <font color=blue><%=rs.PageSize%></font> ������ <font color=red><%=currentpage%></font> ҳ / �� <font color=blue><%=Pcount%></font> ҳ</td>
			<td class=tablebody2 colspan=8 align=right>��ҳ��<%call disppagenum(currentpage,Pcount,"""?page=","""")%>&nbsp;&nbsp;&nbsp;&nbsp;</td>
		</tr>
	<%end if
	rs.close
	set rs=nothing	
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
	sql= "select * from [KeHu] where ZhangHao='"&username&"' and suoding<2"
	set rs=gp_conn.execute(sql)         
	if rs.eof then
		errmess="<li>���û��Ĳ�����"
		call endinfo(2)
	else		
%>
	<tr>
		<th colspan="2" valign=center align=middle height=25>�༭ ����<font color=red><%=rs("ZhangHao")%></font>��������</th>
	</tr>
	<form method=post  action="?action=SaveEdit">
	<tr>
		<td class=tablebody2 width="40%">UID</td>
		<td class=tablebody1 width="60%">&nbsp;<%=rs("id")%><input type=hidden name=Uid value="<%=rs("id")%>"></td>
	</tr>	
	<tr>
		<td class=tablebody2>�ʻ�</td>
		<td class=tablebody1>&nbsp;<font color=blue><%=rs("ZhangHao")%></font><input type=hidden name=username value="<%=rs("ZhangHao")%>"></td>
	</tr>
	<tr>
		<td class=tablebody2>��������</td>
		<td class=tablebody1>&nbsp;<%=rs("KaiHuRiQi")%></td>
	</tr>	
	<tr>
		<td class=tablebody2>�������</td>
		<td class=tablebody1>&nbsp;<%=rs("ZuiHouRiQi")%></td>
	</tr>	
	<tr>
		<td class=tablebody2>״̬</td>
		<td class=tablebody1><input type=radio name=Locked value="0" <%if rs("SuoDing")<>1 then%> checked <%end if%>> ������<input type=radio name=Locked value="1" <%if rs("SuoDing")=1 then%> checked <%end if%>> ���� </td>
	</tr>
	<tr>
		<td class=tablebody2>�ʽ�</td>
		<td class=tablebody1>&nbsp;<input type=text name=Crash value="<%=round(rs("ZiJin"),0)%>"></td>
	</tr>			
	<tr>
		<td class=tablebody2>���ʽ�</td>
		<td class=tablebody1>&nbsp;<input type=text name=TotalFund value="<%=round(rs("ZongZiJin"),0)%>"></td>
	</tr>
	<tr>
		<td class=tablebody2><font color="#333333">�ֹ�����</font></td>
		<td class=tablebody1>&nbsp;<input type=text name=StockNum value="<%=rs("ChiGuZhongLei")%>"></td>
	</tr>
	<tr>
		<td class=tablebody2><font color=gray>��������</font></td>
		<td class=tablebody1>&nbsp;<input type=text name=TodayBuy readonly value="<%=rs("JinRiMaiRu")%>"></td>
	</tr>	
	<tr>
		<td class=tablebody2><font color=gray>��������</font></td>
		<td class=tablebody1>&nbsp;<input type=text name=TodaySale readonly value="<%=rs("JinRiMaiRu")%>"></td>
	</tr>
	<tr>
		<td class=tablebody2><font color=gray>������</font></td> 
		<td class=tablebody1>&nbsp;<input type=text name=TotalBuy readonly value="<%=rs("ZongMaiRu")%>"></td>
	</tr>	
	<tr>
		<td class=tablebody2><font color=gray>������</font></td>
		<td class=tablebody1>&nbsp;<input type=text name=TotalSale readonly value="<%=rs("ZongMaiChu")%>"></td>
	</tr>
	<tr><th colspan="2" align="center"><input type=submit name=submit value=ִ���޸�></th>
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
	sql= "select * from [KeHu] where id="&Uid&""         
	set rs=server.createobject("adodb.recordset")
	rs.open sql,gp_conn,1,3
	rs("ZiJin")=request.form("Crash")
	rs("ZongZiJin")=request.form("TotalFund")
	rs("SuoDing")=request.form("Locked")
	rs("ChiGuZhongLei")=request.form("StockNum")
	rs("JinRiMaiRu")=request.form("TodayBuy")
	rs("JinRiMaiChu")=request.form("TodaySale")
	rs("ZongMaiRu")=request.form("TotalBuy")
	rs("ZongMaiChu")=request.form("TotalSale")		
	rs.update
	rs.close
	sucmess="<li>�û� <font color=blue>"&username&"</font> �����ϸ��³ɹ�"
	rUrl="?"
	call endinfo(2)
end sub		

sub DelUser()
	dim username,NowPrice,rsDaHu,sid,SaleNum,AddMoney,Poundage,Rname,rsGuPiao,mess
	username=trim(replace(request("username"),"'",""))
	if username="" then
		errmess="<li>����Ĳ������������еĲ����Ƿ�"
		call endinfo(2)
		exit sub
	end if	
	sql= "select * from [KeHu] where ZhangHao='"&username&"' and suoding<2"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,gp_conn,1,3
	if rs.eof then
		errmess="<li>���û��Ĳ�����"
		call endinfo(2)
	else		
		sql="select * from [DaHu] where ZhangHao='"&username&"'"
		set rsDaHu=server.createobject("adodb.recordset")
		rsDahu.open sql,gp_conn,1,1
		do while not rsDaHu.eof
			sid=rsDaHu("sid")
			NowPrice=gp_conn.execute("select DangQianJiaGe,QiYe from [GuPiao] where sid="&sid)(0)
			if isnull(NowPrice) then NowPrice=0
			SaleNum=rsDaHu("ChiGuShu")
			AddMoney=SaleNum*NowPrice
			Poundage=AddMoney*Trade_Setting(5)/100
			gp_conn.execute "update [GuPiao] set ChengJiaoLiang=ChengJiaoLiang+"&SaleNum&",ShengYuGuFen=ShengYuGuFen+"&SaleNum&",JiaoYiLiang=JiaoYiLiang+"&SaleNum&",MaiChuBiShu=MaiChuBiShu+1 where sid="&sid
			gp_conn.execute "update [KeHu] set ZiJin=ZiJin+"&AddMoney-Poundage&" where ZhangHao='"&UserName&"'"
			mess="<font color=#FF6666>"&username&"</font><font color=#33CCFF>���ϳ����У���������"&rsDaHu("QiYe")&" "&SaleNum&" ��</font>"
			set rsGuPiao=gp_conn.execute("select JingYingZhe from [GuPiao] where sid="&sid)
			Rname=""
			if lcase(trim(rsGuPiao(0)))=lcase(trim(UserName)) then
				rsGuPiao.close
				set rsGuPiao=gp_conn.execute("select top 1 ChiGuShu,ZhangHao,Uid,QiYe from [DaHu] where sid="&sid&" and Uid>0 and ZhangHao<>'"&UserName&"' order by ChiGuShu desc")
				if not rsGuPiao.eof then
					gp_conn.execute("update [GuPiao] set JingYingZhe='"&rsGuPiao(1)&"',Uid="&rsGuPiao(2)&" where sid="&sid) 
					Rname=rsGuPiao(1)
				else
					gp_conn.execute("update [GuPiao] set JingYingZhe=null,Uid=0 where sid="&sid)
				end if
			end if	
			if Rname<>"" then mess=mess&"<font color=#33CCFF>,<font color=#FF6666>"&Rname&"</font>��øùɵľ�ӪȨ</font>"
			rsGuPiao.close	
			gp_conn.execute "insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )"  
			set rsGuPiao=gp_conn.execute("select Uid from [GuPiao] where sid="&sid)
			if not rsGuPiao.eof then
				gp_conn.execute("update [KeHu] set ZiJin=ZiJin+"&csng(Poundage)&",ZongZiJin=ZongZiJin+"&csng(Poundage)&" where id="&rsGuPiao(0) )
			end if	
			rsGuPiao.close
			gp_conn.execute("update [GupiaoConfig] set TodaySale=TodaySale+1,TodayTotal=TodayTotal+"&AddMoney&"")
			rsDaHu.movenext
		loop
		rsDaHu.close
		gp_conn.execute "delete from [DaHu] where ZhangHao='"&UserName&"'"
		conn.execute("update [user] set userwealth=userwealth+"&rs("ZiJin")&" where username='"&username&"'")
		rs("ZongZiJin")=0
		rs("ZiJin")=0
		rs("SuoDing")=2
		rs("ChiGuZhongLei")=0
		rs("JinRiMaiRu")=0
		rs("JinRiMaiChu")=0
		rs("ZongMaiRu")=0
		rs("ZongMaiChu")=0
		rs.update
		rs.close
		if Gupiao_Setting(12)>0 then
			'ֻ��������Gupiao_Setting(12)���¼�
			gp_conn.execute("delete from RndEvent where id not in(select top "&Gupiao_Setting(12)&" id from RndEvent order by id desc)")
		end if
		sucmess="<li>�û� <font color=blue>"&username&"</font> ���Ѿ���ɾ�������й�Ʊ�ּ��׳�"
		rUrl="?"
		call endinfo(2)
	end if
end sub		
%>