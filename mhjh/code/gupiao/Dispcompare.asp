<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html><head>
<META http-equiv=Content-Type content=text/html; charset=gb2312>
<meta name=keywords content="快乐江湖股票市场">
<title><%=Gupiao_Setting(5)%>-上市公司资讯</title>
<!--#include file="css.asp"-->
</head>
<body topmargin=5 leftmargin=0 oncontextmenu=self.event.returnValue=false>
<%
	dim Sid,rsu,TongJi,TongJiF,TotalNum,UserNum
	dim GPScale(),Tnum
	const VHeight=200
	const VWidth=150
	stats="上市公司资讯"
	Sid=request("Sid")
	if not isnumeric(Sid) then
		errmess="<li>错误的参数，请指定正确的股票"
		call endinfo(1)
	else
		set rs=conn.execute("select * from [股票] where sid="&Sid)
		if rs.eof and rs.bof then
			errmess="<li>没有找到指定的上市公司,可能该上市公司已经倒闭"
			call endinfo(1)			
		else
			TongJi=rs("TongJi")
			if TongJi="" or isnull(TongJi) then TongJi="0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"
			TongJiF=TongJi
			TongJi=split(TongJi,"|")
			TongJiF=split(TongJiF,"|")
			
			TotalNum=conn.execute("select sum(持股数) from [大户] where sid="&sid&" and uid>0")(0)
			UserNum=conn.execute("select count(*) from [大户] where sid="&sid&" and uid>0")(0)
			if TotalNum=0 then TotalNum=1
			
			call FormatValue(VHeight)
			Tnum=StockScale(VWidth,UserNum,TotalNum)			
%>
		<table  width="97%" border=0 cellspacing=0 cellpadding=0 align=center bgcolor="#E4E8EF">
			<TR>
				<TD BACKGROUND="Images/topbg.gif" height=9 colspan=3>
			</TD>
			</TR>
			<tr>
				<td valign=center align=middle height=23 background="Images/Header.gif"><font class=big>上市公司 <font color=red><%=rs("企业")%></font> 资讯</font></td>
			</tr>
		</table>
		<br>
		<table width="97%" border="0" cellspacing="1" cellpadding="3" align="center" style="border: 1px #6595D6 solid; background-color: #FFFFFF;">
			<TR valign="middle"> 
				<td class=forumRow bgcolor="#E4E8EF" align="center" width="150" rowspan="8"><%if rs("LtdImg")<>"" and (not isnull(rs("LtdImg"))) then%><a href="Exchange.asp?sid=<%=rs("sid")%>" title="进入交易页面"><img border="0" src="<%=rs("LtdImg")%>" width="150"></a><%else%><font color=gray>无图片</font><%end if%></td>
				<td class=forumRow height="20"><b>上市公司</b></td>
				<td class=forumRow>&nbsp;<a href="Exchange.asp?sid=<%=rs("sid")%>" class="cblue" title="进入交易页面"><%=rs("企业")%></a></td>
				<td class=forumRow height="20"><b>经营者</b></td>
				<td class=forumRow>&nbsp;<%=rs("经营者")%></td>				
			</TR>
			<TR valign="middle"> 
				<td class=forumRow height="20"><b>股票代号</b></td>
				<td class=forumRow>&nbsp;<%=rs("sid")%></td>
				<td class=forumRow height="20"><b>上市日期</b></td>
				<td class=forumRow>&nbsp;<%=rs("日期")%></td>				
			</TR>			
			<TR valign="middle"> 
				<td class=forumRow height="20"><b>总股份</b></td>
				<td class=forumRow>&nbsp;<%=rs("总股份")%> 股</td> 
				<td class=forumRow height="20"><b>剩余股份</b></td> 
				<td class=forumRow>&nbsp;<%=rs("剩余股份")%> 股</td>
			</TR>
			<TR valign="middle"> 
				<td class=forumRow height="20"><b>交易量</b></td>
				<td class=forumRow>&nbsp;<%=rs("交易量")%> 股</td> 
				<td class=forumRow height="20"><b>成交量</b></td> 
				<td class=forumRow>&nbsp;<%=rs("成交量")%> 股</td>
			</TR>
			<TR valign="middle"> 
				<td class=forumRow height="20"><b>买入笔数</b></td>
				<td class=forumRow>&nbsp;<%=rs("买入笔数")%> 笔</td> 
				<td class=forumRow height="20"><b>卖出笔数</b></td> 
				<td class=forumRow>&nbsp;<%=rs("卖出笔数")%> 笔</td> 
			</TR>
			<TR valign="middle"> 
				<td class=forumRow height="20"><b>开盘价格</b></td>
				<td class=forumRow>&nbsp;<%=formatcurrency(rs("开盘价格"),2,-1)%></td>
				<td class=forumRow height="20"><b>当前价格</b></td>
				<td class=forumRow>&nbsp;<%=formatcurrency(rs("当前价格"),2,-1)%></td>				
			</TR>			
			<TR valign="middle"> 
				<td class=forumRow height="20"><b>今日涨跌</b></td>
				<td class=forumRow>&nbsp;<%=formatcurrency(rs("当前价格")-rs("开盘价格"),2,true)%></td>  
				<td class=forumRow height="20"><b>今日涨跌幅</b></td>   
				<td class=forumRow>&nbsp;<%=formatnumber( (rs("当前价格")-rs("开盘价格"))*100/rs("开盘价格"),2,true)%> %</td>   
			</TR>	 		
			<TR valign="middle"> 
				<td class=forumRow height="20"><b>今日指数</b></td>
				<td class=forumRow>&nbsp;<%=formatnumber(rs("TodayWave"),2,-1)%> 点</td> 
				<td class=forumRow height="20"><b>综合指数</b></td>  
				<td class=forumRow>&nbsp;<%=formatnumber(rs("TotalWave"),2,-1)%> 点</td>  
			</TR>									
			<TR valign="middle"> 
				<td class=forumRow height="*" align="center"><b>公司简介</b></td>
				<td class=forumRow colspan="4"><%=rs("Explain")%></td>
			</TR>
		</table> 
		<br>
		<table cellspacing=1 cellpadding=0 width="97%" align=center border=0>
		<tr> 
			<td width="55%"> 	
				<table  width="100%" border=0 cellspacing=1 cellpadding=3 align=center style="border: 1px #6595D6 solid; background-color: #FFFFFF;">													
					<tr>
						<th height=25 align="center" colspan=9 class="admin">最近30天走势图</th>
					</tr> 
					<tr> 
						
						<td class="forumRow" valign="bottom" height="<%=VHeight+20%>" nowrap>
						<%
						for i=0 to ubound(tongji)%>
							<img src="images/bar.gif" width=7 height="<%=TongJiF(i)+5%>" alt="当日收盘价：<%=TongJi(i)%> 元">
						<%next%>				
						</td>
					</tr>
				</table>
			</td>
			<td width="3"></td>
			<td>
				<table  width="100%" border=0 cellspacing=1 cellpadding=3 align=center style="border: 1px #6595D6 solid; background-color: #FFFFFF;">													
					<tr>
						<th height=25 align="center" colspan=9 class="admin">十大股东持股比例图</th>
					</tr> 
					<tr><td class="forumRow" valign="top" height="<%=VHeight+20%>">
					<table  width="100%" height='100%' border=0 cellspacing=0 cellpadding=3 align=center>
						<%if UserNum<=0 then
							response.Write "<tr valign=middle align=center><td><font color=gray>暂时没有人购买该股票</td></tr>"
						else
						for i=0 to Tnum%>					
						<tr>
							<td width="90" nowrap><%=GPScale(i,0)%></td>
							<td nowrap><img src="images/baru.gif" height=10 width="<%=GPScale(i,3)+5%>" alt="持股数：<%=GPScale(i,1)%>股 持股百分比：<%=formatnumber(GPScale(i,2)*100,2,-1)%>%"></td>
						</tr>
						<%next
						end if
						%>
					</table>
					</td></tr>
				</table>			
			</td>
		</tr>
		</table>				 
		<br>	
		<table  width="97%" border=0 cellspacing=1 cellpadding=3 align=center style="border: 1px #6595D6 solid; background-color: #FFFFFF;">													

		  <TR align=center valign="middle" > 
			  <th height=25>股东</TD>
			  <th>持股数量</TD>
			  <th>持股比例</TD>
			  <th>当前价格</TD>
			  <th>持股价格</TD>
			  <th>持股现值</TD>  
			  <th>持股成本</TD> 			  
			  <th>总盈亏</TD>   
			  <th>盈亏率</TD>   
		  </TR> 			
<%
		set rsu=conn.execute("select * from [大户] where sid="&sid&" and uid>0 order by 持股数 desc") 
		if rsu.eof and rsu.bof then 
			response.Write "<tr><td colspan=9><font color=gray>暂时没有人购买这个股票</td></tr>"
		else
			dim GupiaoValue,GupiaoCost
			do while not rsu.eof 
				GupiaoValue=rs("当前价格")*rsu("持股数")		'持股现值
				GupiaoCost =rsu("平均价格")*rsu("持股数")		'持股成本
%>
		  <TR align=center valign="middle" > 
			  <td class="bg1" height=25><a href='dispu.asp?Uid=<%=rsu("uid")%>' <%if membername=rsu("帐号") then%> class="cblue" <%end if%>><%=rsu("帐号")%></a></TD>
			  <td class="bg1"><%=rsu("持股数")%></TD>
			  <td class="bg1"><%=formatnumber(rsu("持股数")*100/TotalNum,2,-1)%> %</TD>
			  <td class="bg1"><%=formatcurrency(rs("当前价格"),2,-1)%></TD>
			  <td class="bg1"><%=formatcurrency(rsu("平均价格"),2,-1)%></TD>
			  <td class="bg1"><%=formatcurrency(GupiaoValue,2,-1)%></TD>  
			  <td class="bg1"><%=formatcurrency(GupiaoCost,2,-1)%></TD>
			  <td class="bg1"><%=formatcurrency(GupiaoValue-GupiaoCost,2,-1)%></TD>   
			  <td class="bg1"><%if GupiaoCost=0 then%>0.00<%else%><%=formatnumber((GupiaoValue-GupiaoCost)*100/GupiaoCost,2,-1)%><%end if%> %</TD>  
		  </TR>
<%				
				rsu.movenext
			loop
		end if
%>			
		</table>		
		<br>	
		<table  width="97%" border=0 cellspacing=0 cellpadding=0 align=center bgcolor="#E4E8EF">													
			<tr>
				<td align=center background="Images/footer.gif" height=30><a href="Exchange.asp?sid=<%=rs("sid")%>" class="cblue" title="进入交易">[交易]</a>&nbsp;&nbsp;<a href="stock.asp">[返回股市]</a></td>
			</tr>
		</table>  
<%
		end if
		rs.close
	end if
	call GPOnline()	 	'在线股民
	CloseDatabase		'关闭数据库
%>
</BODY>
</HTML>
<%
function FormatValue(StdValue)
	dim Max,ii
	'找出最大值
	Max=TongJi(0)
	for ii=1 to ubound(TongJi)
		if csng(Max)<csng(TongJi(ii)) then Max=TongJi(ii)	
	next
	if max=0 then max=1
	'格式化数据   
	for ii=0 to ubound(TongJi)
	 	TongJiF(ii)=int(TongJi(ii)*StdValue/Max)
	next  
end function   

'10大股东持股比例,StdWidth-标准长度，TopNum<10 ,TotalN-所有股东持股数目
function StockScale(StdWidth,TopNum,TotalN) 
	dim other,num
	other=false
	StockScale=0
	num=TopNum 
	if num<=0 then 
		exit function
	elseif num>10 then
		num=9
		other=true
	end if	
	Redim GPScale(num,4)
	Dim rst,ii
	set rst=conn.execute("select top "&num&" 帐号,持股数 from [大户] where sid="&sid&" and uid>0 order by 持股数 desc") 
	for ii=0 to num-1
		GPScale(ii,0)=rst(0)				'用户名
		GPScale(ii,1)=rst(1)				'持股数
		GPScale(ii,2)=rst(1)/TotalN			'持股百分比	 
		GPScale(ii,3)=int(GPScale(ii,2)*StdWidth)	'格式化长度  
		StockScale=StockScale+GPScale(ii,1)		'前9个股东持股数之和 
		rst.movenext 
	next 
	if other then   
		GPScale(9,0)="<font color=gray>其他股东</font>"    
		GPScale(9,1)=TotalN-StockScale
		GPScale(9,2)=GPScale(9,1)/TotalN
		GPScale(9,3)=int(GPScale(ii,2)*StdWidth)
	end if
	StockScale=iif(Topnum>10,9,Topnum-1)
end function
%>