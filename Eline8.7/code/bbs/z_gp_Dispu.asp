<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_gp_conn.asp"-->
<!--#include file="z_gp_Const.asp"-->
<%
stats="股民持股"
call nav()
call head_var(0,0,GuPiao_Setting(5),"z_gp_gupiao.asp")
dim Uid,UserName,UserCash,Userfund
dim StockType,TodayBuy,TodaySale,TotalBuy,TotalSale
Uid=request("uid")
if not isnumeric(Uid) then
	errmess="<li>错误的参数，请指定正确的用户ID"
	call endinfo(2)
else
	set rs=gp_conn.execute("select * from [KeHu] where id="&Uid&" and suoding<2")
	if rs.eof and rs.bof then
		errmess="<li>指定的用户没有在股市开户"
		errmess=errmess+"<br><li>指定的用户不存在，请确认传递的参数是否正确"
		call endinfo(2)
	else
		UserName=rs("ZhangHao")
		UserCash=rs("ZiJin")				
		Userfund=rs("ZongZiJin")
		StockType=rs("ChiGuZhongLei")
		TodayBuy=rs("JinRiMaiRu")
		TodaySale=rs("JinRiMaiChu")
		TotalBuy=rs("ZongMaiRu")
		TotalSale=rs("ZongMaiChu")
		call main()
	end if
end if
call activeonline()
call footer()
'=====================================	
sub main()%>
	<table class=tableborder1 cellspacing=1 cellpadding=3 align=center border=0 width="<%=Forum_body(12)%>">
		<tr>
			<th colspan=9 valign=center align=middle height=25>股民 <font color=red><%=UserName%></font> 持股一览</th>
		</tr>
	  <TR align=center valign="middle" > 
		  <td class=tablebody2 align=center><b>股票名称</b></Td>
		  <td class=tablebody2 align=center><b>拥有数量</b></Td>
		  <td class=tablebody2 align=center><b>持股成本</b></Td>
		  <td class=tablebody2 align=center><b>当前价格</b></Td>
		  <td class=tablebody2 align=center><b>持股价格</b></Td>
		  <td class=tablebody2 align=center><b>持股现值</b></Td>
		  <td class=tablebody2 align=center><b>每股盈亏</b></Td> 
		  <td class=tablebody2 align=center><b>总盈亏</b></Td>
			<td class=tablebody2 align=center><b>盈亏率</b></Td> 
		</tr>
		<%Dim MyGupiaoNum,MyGupiaoValue,MyGupiaoCost 		'个人持股总数，个人股票市值,个人股票成本
		Dim TotalGranRate  '总盈亏率
		MyGupiaoNum=0 :	MyGupiaoValue=0 :	MyGupiaoCost=0:		TotalGranRate=0
		sql="select d.*,g.DangQianJiaGe from [DaHu] d inner join [GuPiao] g on d.sid=g.sid where d.uid="&Uid&" order by d.[ChiGuShu] desc" 
		set rs=gp_conn.execute(sql)
		if rs.eof and rs.bof then
			RESPONSE.Write "<tr><td colspan=9 class=tablebody1>没有股票</tr>"
		else
			Dim rst,CurrentPrice,AgvInPrice		'某股票当前价格 ,某股票平均买入价格
			do while not rs.EOF   
				CurrentPrice=rs("DangQianJiaGe")
				AgvInPrice=iif(rs("PingJunJiaGe")=0,rs("MaiRuJiaGe"),rs("PingJunJiaGe"))
				MyGupiaoNum=MyGupiaoNum+rs("ChiGuShu")
				MyGupiaoValue=MyGupiaoValue+rs("ChiGuShu")*CurrentPrice
				MyGupiaoCost=MyGupiaoCost+rs("ChiGuShu")*AgvInPrice%>
				<TR align="center" height=25> 
					<TD class=tablebody1><a href="z_gp_Exchange.asp?sid=<%=rs("sid")%>"><%=rs("QiYe")%></a></TD>
					<TD class=tablebody1><%=rs("ChiGuShu")%></TD>
					<TD class=tablebody1 align=right><%=formatnumber(AgvInPrice*rs("ChiGuShu"),2,true)%>&nbsp;</TD>
					<TD class=tablebody1 align=right><%=formatnumber(CurrentPrice,2,true)%>&nbsp;</TD>
					<TD class=tablebody1 align=right><%=formatnumber(AgvInPrice,2,true)%>&nbsp;</TD>
					<TD class=tablebody1 align=right><%=formatnumber(CurrentPrice*rs("ChiGuShu"),2,true)%>&nbsp;</TD>
					<TD class=tablebody1 align=right><%
						dim bj,TotalGain,GranRate		'单股盈亏,单个股票总盈亏,单个股票总盈亏率
						bj=CurrentPrice-AgvInPrice
						bj=round(bj,2)
						if bj>0 then
							response.write "<font color=#FF0000>↑ "&formatnumber(bj,2,true)&"</font>"
						elseif bj<0 then
							response.write "<font color=#0000FF>↓ "&formatnumber(abs(bj),2,True)&"</font>"
						elseif bj=0 then
							response.write "<font color=#000000>→ 0.00</font>"
						end if
					%>&nbsp;</TD>
					<TD class=tablebody1 align=right><%
						TotalGain=bj*rs("ChiGuShu")
						GranRate=TotalGain*100/MyGupiaoCost
						if TotalGain>0 then 
							response.Write "<font color=#FF0000>↑ "&formatnumber(TotalGain,2,True)&"</font>&nbsp;</TD>" 
							response.Write "<TD class=tablebody1 align=right><font color=#FF0000>↑"&formatnumber(GranRate,2,True)&" %</font>"
						elseif TotalGain<0 then
							response.Write "<font color=#0000FF>↓ "&formatnumber(-TotalGain,2,True)&"</font>&nbsp;</TD>"
							response.Write "<TD class=tablebody1 align=right><font color=#0000FF>↓"&formatnumber(-GranRate,2,True)&" %</font>" 
						else
							response.Write "<font color=#000000>→ 0.00</font>&nbsp;</TD>" 
							response.Write "<TD class=tablebody1 align=right><font color=#000000>→ 0.00 %</font>"	
						end if
					%>&nbsp;</TD>
				</tr>
				<%rs.MoveNext          
			Loop
			rs.close
		end if 
		if MyGupiaoCost=0 then
			TotalGranRate=0
		else	
			TotalGranRate=(MyGupiaoValue-MyGupiaoCost)*100/MyGupiaoCost
		end if%>	
		<TR align=center valign="middle" height=25> 
			<Td class=tablebody2>合计</Td> 
			<Td class=tablebody2><%=formatnumber(MyGupiaoNum,0,true)%></Td>
			<Td class=tablebody2 align=right><%=formatnumber(MyGupiaoCost,2,true)%>&nbsp;</Td>
			<td class=tablebody2>-</Td>
			<Td class=tablebody2>-</Td>
			<Td class=tablebody2 align=right><%=formatnumber(MyGupiaoValue,2,true)%>&nbsp;</Td>
			<Td class=tablebody2>-</Td>
			<Td class=tablebody2 align=right><%if MyGupiaoValue-MyGupiaoCost>0 then%>↑ <%else%>↓ <%end if%><%=formatnumber(abs(MyGupiaoValue-MyGupiaoCost),2,true)%>&nbsp;</Td> 
			<Td class=tablebody2 align=right><%if TotalGranRate>0 then%>↑ <%else%>↓ <%end if%><%=formatnumber(TotalGranRate,2,true)%> %&nbsp;</Td>
		</TR>
	</TABLE>
	<br>
	<table class=tableborder1 cellspacing=1 cellpadding=3 align=center border=0 width="<%=Forum_body(12)%>">
		<TR align=center valign="middle" height="25" > 
			<th>股民</th>
			<th>总资产</th>
			<th>资金</th>
			<th>持股总价值</th>
			<th>持股种类</th>
			<th>今日买入</th>
			<th>今日卖出</th>
			<th>总买入</th>
			<th>总卖出</th></tr> 
		</TR>
		<tr align='center' height='22'>
			<td class=tablebody1><%=Username%></td>
			<td class=tablebody1 align=right><%=formatnumber(Userfund,0,true)%>&nbsp;</td>
			<td class=tablebody1 align=right><%=formatnumber(UserCash,0,true)%>&nbsp;</td>
			<td class=tablebody1 align=right><%=formatnumber(MyGupiaoValue,0,true)%>&nbsp;</td>
			<td class=tablebody1><%=StockType%></td>
			<td class=tablebody1 align=right><%=TodayBuy%>&nbsp;</td>
			<td class=tablebody1 align=right><%=TodaySale%>&nbsp;</td>
			<td class=tablebody1 align=right><%=TotalBuy%>&nbsp;</td>
			<td class=tablebody1 align=right><%=TotalSale%>&nbsp;</td>
		</tr>
		<tr>
			<td class=tablebody2 align=center colspan=9 height=25><a href="z_gp_TopUser.asp"><b>[返回排行榜]</b></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="z_gp_Gupiao.asp"><b>[返回股市]</b></a></td>
		</tr>
	</table>  
<%end sub
%>