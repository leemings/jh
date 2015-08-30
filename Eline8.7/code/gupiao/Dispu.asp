<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html><head>
<META http-equiv=Content-Type content=text/html; charset=gb2312>
<meta name=keywords content="『快乐江湖』股票市场">
<title><%=Gupiao_Setting(5)%>-股民持股一览表</title>
<!--#include file="css.asp"-->
</head>
<body topmargin=5 leftmargin=0 oncontextmenu=self.event.returnValue=false>
<%
	dim Uid,UserName,UserCash,Userfund
	dim StockType,TodayBuy,TodaySale,TotalBuy,TotalSale
	Uid=request("uid")
	if not isnumeric(Uid) then
		errmess="<li>错误的参数，请指定正确的用户ID"
		call endinfo(2)
	else
		set rs=conn.execute("select * from [客户] where id="&Uid)
		if rs.eof and rs.bof then
			errmess="<li>指定的用户没有在股市开户"
			errmess=errmess+"<br><li>指定的用户不存在，请确认传递的参数是否正确"
			call endinfo(2)
		else
			UserName=rs("帐号")
			UserCash=rs("资金")				
			Userfund=rs("总资金")
			StockType=rs("持股种类")
			TodayBuy=rs("今日买入")
			TodaySale=rs("今日卖出")
			TotalBuy=rs("总买入")
			TotalSale=rs("总卖出")
			call main()
		end if
	end if
	CloseDatabase		'关闭数据库
'=====================================	
sub main()		
%>
<table  width="97%" border=0 cellspacing=0 cellpadding=0 align=center bgcolor="#E4E8EF">
	<TR>
		<TD BACKGROUND="Images/topbg.gif" height=9 colspan=3>
	</TD>
	</TR>
	<tr>
		<td valign=center align=middle height=23 background="Images/Header.gif"><font class=big>股民 <font color=red><%=UserName%></font> 持股一览</font></td>
	</tr>
</table>
<br>
<table border="0" width="97%" bgcolor="#666666" cellspacing="1" cellpadding="3" bordercolorlight="#000000" bordercolordark="#6E6E6E"  align="center">
  <TR align=center bgColor=darkblue valign="middle" > 
	  <td background='Images/lan.gif' height=25><font color="#FFFFFF">代号</TD>
	  <td background='Images/lan.gif' ><font color="#FFFFFF">股票名称</TD>
	  <td background='Images/lan.gif' ><font color="#FFFFFF">拥有数量</TD>
	  <td background='Images/lan.gif' ><font color="#FFFFFF">持股成本</TD>
	  <td background='Images/lan.gif' ><font color="#FFFFFF">当前价格</TD>
	  <td background='Images/lan.gif' ><font color="#FFFFFF">持股价格</TD>
	  <td background='Images/lan.gif' ><font color="#FFFFFF">持股现值</TD>
	  <td background='Images/lan.gif' ><font color="#FFFFFF">每股盈亏</TD> 
	  <td background='Images/lan.gif' ><font color="#FFFFFF">总盈亏</TD>
      <td background='Images/lan.gif' ><font color="#FFFFFF">盈亏率</TD> 
  </TR>
<%
	Dim MyGupiaoNum,MyGupiaoValue,MyGupiaoCost 		'个人持股总数，个人股票市值,个人股票成本
	Dim TotalGranRate  '总盈亏率
	MyGupiaoNum=0 :	MyGupiaoValue=0 :	MyGupiaoCost=0:		TotalGranRate=0
	sql="select d.*,g.当前价格 from [大户] d inner join [股票] g on d.sid=g.sid where d.uid="&Uid&" order by d.[持股数] desc" 
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		RESPONSE.Write "<tr><td colspan=10 bgcolor='#8d9aca'>没有股票</tr>"
	else
		Dim rst,CurrentPrice,AgvInPrice		'某股票当前价格 ,某股票平均买入价格
		do while not rs.EOF   
			CurrentPrice=rs("当前价格")
			AgvInPrice=iif(rs("平均价格")=0,rs("买入价格"),rs("平均价格"))
			MyGupiaoNum=MyGupiaoNum+rs("持股数")
			MyGupiaoValue=MyGupiaoValue+rs("持股数")*CurrentPrice
			MyGupiaoCost=MyGupiaoCost+rs("持股数")*AgvInPrice	
%>
		<TR align="center" bgColor="#8d9aca"> 
			<TD><a href="exchange.asp?sid=<%=rs("sid")%>"><font color="#FFFF00"><%=rs("sid")%></font></a></TD>
			<TD><a href="Dispcompare.asp?sid=<%=rs("sid")%>"><font color="#FFFF00"><%=rs("企业")%></font></a></TD>
			<TD><font color="#FFFF00"><%=rs("持股数")%></font></TD>
			<TD><font color="#FFFF00">￥<%=formatnumber(AgvInPrice*rs("持股数"),3,True)%></font></TD>
			<TD><font color="#FFFF00">￥<%=formatnumber(CurrentPrice,3,True)%></font></TD>
			<TD><font color="#FFFF00">￥<%=formatnumber(AgvInPrice,3,True)%></font></TD>
			<TD><font color="#FFFF00">￥<%=formatnumber(CurrentPrice*rs("持股数"),3,True)%></font></TD>
			<TD>
<%
			dim bj,TotalGain,GranRate		'单股盈亏,单个股票总盈亏,单个股票总盈亏率
			bj=CurrentPrice-AgvInPrice
			bj=formatnumber(bj,3,True)
			if bj>0 then
				response.write "<font color=#FF0000>↑"&bj&" 元</font>"
			elseif bj<0 then
				response.write "<font color=#00FF00>↓"&formatnumber(abs(bj),3,True)&" 元</font>"
			elseif bj=0 then
				response.write "<font color=#FFFFff>→ 0.000 元</font>"
			end if
			response.Write "</TD><TD>"
			
			TotalGain=bj*rs("持股数")
			GranRate=TotalGain*100/MyGupiaoCost
			if TotalGain>0 then 
				response.Write "<font color=#FF0000>↑"&formatnumber(TotalGain,3,True)&" 元</font></TD>" 
				response.Write "<TD><font color=#FF0000>↑"&formatnumber(GranRate,2,True)&" %</font></TD></tr>"
			elseif TotalGain<0 then
				response.Write "<font color=#00FF00>↓"&formatnumber(-TotalGain,3,True)&" 元</font></TD>"
				response.Write "<TD><font color=#00FF00>↓"&formatnumber(-GranRate,2,True)&" %</font></TD></tr>" 
			else
				response.Write "<font color=#FFFFff>→ 0.00 元</font></TD>" 
				response.Write "<TD><font color=#FFFFff>→ 0.00 %</font></TD></tr>"	
			end if	
			rs.MoveNext          
		Loop
		rs.close
	end if 
	if MyGupiaoCost=0 then
		TotalGranRate=0
	else	
		TotalGranRate=(MyGupiaoValue-MyGupiaoCost)*100/MyGupiaoCost
	end if	
%>	
		<TR align=center bgColor="#99CCFF" valign="middle"> 
			<TD background="Images/title.gif" height=20 colspan="2"><font color="#3366CC">合计</font></TD> 
			<TD background="Images/title.gif"><font color="#3366CC"><%=formatnumber(MyGupiaoNum,0,-1)%> 股</font></TD>
			<TD background="Images/title.gif"><font color="#3366CC"><%=formatnumber(MyGupiaoCost,2,-1)%> 元</font></TD>
			<td background="Images/title.gif"><font color="#3366CC">-</font></TD>
			<TD background="Images/title.gif"><font color="#3366CC">-</font></TD>
			<TD background="Images/title.gif"><font color="#3366CC"><%=formatnumber(MyGupiaoValue,2,-1)%> 元</font></TD>
			<TD background="Images/title.gif"><font color="#3366CC">-</font></TD>
			<TD background="Images/title.gif"><font color="#3366CC"><%if MyGupiaoValue-MyGupiaoCost>0 then%>+<%end if%><%=formatnumber(MyGupiaoValue-MyGupiaoCost,2,-1)%></font></TD> 
			<TD background="Images/title.gif"><font color="#3366CC"><%if TotalGranRate>0 then%>+<%end if%><%=formatnumber(TotalGranRate,2,-1)%> %</font></TD>
		</TR>
	</TABLE>
<%
%>	
</table>
<br>
<table align=center bgcolor="#0066CC" border=0 cellspacing=1 cellpadding="3" width="97%">
	<TR align=center bgColor=darkblue valign="middle" height="25" > 
		<td background='Images/lan.gif'><B><font color='FFFFFF'>股民</font></B></td>
		<td background='Images/lan.gif'><B><font color='FFFFFF'>总资产</font></B></td>
		<td background='Images/lan.gif'><B><font color='FFFFFF'>资金</font></B></td>
		<td background='Images/lan.gif'><B><font color='FFFFFF'>持股总价值</font></B></td>
		<td background='Images/lan.gif'><B><font color='FFFFFF'>持股种类</font></B></td>
		<td background='Images/lan.gif'><B><font color='FFFFFF'>今日买入</font></B></td>
		<td background='Images/lan.gif'><B><font color='FFFFFF'>今日卖出</font></B></td>
		<td background='Images/lan.gif'><B><font color='FFFFFF'>总买入</font></B></td>
		<td background='Images/lan.gif'><B><font color='FFFFFF'>总卖出</font></B></td></tr> 
	</TR>
	<tr bgcolor='#E4E6EF'align='center' height='22'>
		<td><%=Username%></td>
		<td><%=formatnumber(Userfund,0,true)%> 元</td>
		<td><%=formatnumber(UserCash,0,true)%> 元</td>
		<td><%=formatnumber(MyGupiaoValue,0,true)%> 元</td>
		<td><%=StockType%></td>
		<td><%=TodayBuy%></td>
		<td><%=TodaySale%></td>
		<td><%=TotalBuy%></td>
		<td><%=TotalSale%></td>
	</tr>
	<tr>
		<td align=center colspan=9 background="Images/footer.gif" height=30><a href="TopUser.asp">[返回排行榜]</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="stock.asp">[返回股市]</a></td>
	</tr>
</table>  
</BODY>
</HTML>
<%
end sub
%>