<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_gp_conn.asp"-->
<!--#include file="z_gp_Const.asp"-->
<%
stats="上市公司资讯"
call nav()
call head_var(0,0,GuPiao_Setting(5),"z_gp_gupiao.asp")
if founderr then
	call dvbbs_error()
else
	dim Sid,rsu,TongJi,TongJiF,TotalNum,UserNum
	dim GPScale(),Tnum,ExplainSplit
	const VHeight=200
	const VWidth=150
	Sid=request("Sid")
	if not isnumeric(Sid) then
		errmess="<li>错误的参数，请指定正确的股票"
		call endinfo(1)
	else
		set rs=gp_conn.execute("select * from [GuPiao] where sid="&Sid)
		if rs.eof and rs.bof then
			errmess="<li>没有找到指定的上市公司,可能该上市公司已经倒闭"
			call endinfo(1)			
		else
			TongJi=rs("TongJi")
			if TongJi="" or isnull(TongJi) then TongJi="0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"
			TongJiF=TongJi
			TongJi=split(TongJi,"|")
			TongJiF=split(TongJiF,"|")
			
			TotalNum=gp_conn.execute("select sum(ChiGuShu) from [DaHu] where sid="&sid&" and uid>0")(0)
			UserNum=gp_conn.execute("select count(*) from [DaHu] where sid="&sid&" and uid>0")(0)
			if UserNum>10 then UserNum=10
			if TotalNum=0 then TotalNum=1
			
			call FormatValue(VHeight)
			Tnum=StockScale(VWidth,UserNum,TotalNum)%>
			<table class=tableborder1 cellspacing=1 cellpadding=3 align=center border=0 width="<%=Forum_body(12)%>">
				<TR>
					<th colspan=5 valign=center align=middle height=25>上市公司 <font color=red><%=rs("QiYe")%></font> 资讯</th>
				</tr>
				<TR valign="middle"> 
					<td class=tablebody1 align="center" width="150" rowspan="8"><%if rs("LtdImg")<>"" and (not isnull(rs("LtdImg"))) then%><a href="z_gp_Exchange.asp?sid=<%=rs("sid")%>" title="进入交易页面"><img border="0" src="<%=rs("LtdImg")%>" width="150"></a><%else%><font color=gray>无图片</font><%end if%></td>
					<td class=tablebody2 height="20"><b>上市公司</b></td>
					<td class=tablebody1>&nbsp;<a href="z_gp_Exchange.asp?sid=<%=rs("sid")%>" class="cblue" title="进入交易页面"><%=rs("QiYe")%></a></td>
					<td class=tablebody2 height="20"><b>经营者</b></td>
					<td class=tablebody1>&nbsp;<%=rs("JingYingZhe")%></td>				
				</TR>
				<TR valign="middle"> 
					<td class=tablebody2 height="20"><b>股票代号</b></td>
					<td class=tablebody1>&nbsp;<%=rs("sid")%></td>
					<td class=tablebody2 height="20"><b>上市日期</b></td>
					<td class=tablebody1>&nbsp;<%=rs("RiQi")%></td>				
				</TR>			
				<TR valign="middle"> 
					<td class=tablebody2 height="20"><b>总股份</b></td>
					<td class=tablebody1>&nbsp;<%=rs("ZongGuFen")%> 股</td> 
					<td class=tablebody2 height="20"><b>剩余股份</b></td> 
					<td class=tablebody1>&nbsp;<%=rs("ShengYuGuFen")%> 股</td>
				</TR>
				<TR valign="middle"> 
					<td class=tablebody2 height="20"><b>交易量</b></td>
					<td class=tablebody1>&nbsp;<%=rs("JiaoYiLiang")%> 股</td> 
					<td class=tablebody2 height="20"><b>成交量</b></td> 
					<td class=tablebody1>&nbsp;<%=rs("ChengJiaoLiang")%> 股</td>
				</TR>
				<TR valign="middle"> 
					<td class=tablebody2 height="20"><b>买入笔数</b></td>
					<td class=tablebody1>&nbsp;<%=rs("MaiRuBiShu")%> 笔</td> 
					<td class=tablebody2 height="20"><b>卖出笔数</b></td> 
					<td class=tablebody1>&nbsp;<%=rs("MaiChuBiShu")%> 笔</td> 
				</TR>
				<TR valign="middle"> 
					<td class=tablebody2 height="20"><b>开盘价格</b></td>
					<td class=tablebody1>&nbsp;<%=formatnumber(rs("KaiPanJiaGe"),2,true)%></td>
					<td class=tablebody2 height="20"><b>当前价格</b></td>
					<td class=tablebody1>&nbsp;<%=formatnumber(rs("DangQianJiaGe"),2,true)%></td>				
				</TR>			
				<TR valign="middle"> 
					<td class=tablebody2 height="20"><b>今日涨跌</b></td>
					<td class=tablebody1>&nbsp;<%=formatnumber(rs("DangQianJiaGe")-rs("KaiPanJiaGe"),2,true)%></td> 
					<td class=tablebody2 height="20"><b>今日涨跌幅</b></td>  
					<td class=tablebody1>&nbsp;<%=formatnumber(round((rs("DangQianJiaGe")-rs("KaiPanJiaGe"))/rs("KaiPanJiaGe"),4)*100,2,true)%> %</td>  
				</TR>			
				<TR valign="middle"> 
					<td class=tablebody2 height="20"><b>今日指数</b></td>
					<td class=tablebody1>&nbsp;<%=formatnumber(rs("TodayWave"),2,true)%> 点</td> 
					<td class=tablebody2 height="20"><b>综合指数</b></td>  
					<td class=tablebody1>&nbsp;<%=formatnumber(rs("TotalWave"),2,true)%> 点</td>  
				</TR>									
				<TR valign="middle"> 
					<td class=tablebody2 height="*" align="center"><b>公司简介</b></td>
					<td class=tablebody1 colspan="4"><%
					ExplainSplit=split(rs("Explain"),"|")%>
					<%=ExplainSplit(0)%></td>
				</TR>
			</table>
			<br>
			<table cellspacing=1 cellpadding=0 border="0" class=tableborder1 align="center">
				<tr> 
					<td width="55%" class=tablebody2> 	
						<table cellspacing=1 cellpadding=0 border=0 align=center width=100%>
							<tr>
								<th height=25 align="center" colspan=3>最近30天走势图</th>
							</tr> 
							<tr> 
								<td class=tablebody1 colspan="3" valign="bottom" height="<%=VHeight+20%>" nowrap align=center>&nbsp;<%
									for i=0 to ubound(tongji)
										%><img src="images/img_gupiao/bar.gif" width=7 height="<%=TongJiF(i)+5%>" alt="当日收盘价：<%=TongJi(i)%> 元">&nbsp;<%
									next
								%>&nbsp;</td>
							</tr>
						</table>
					</td>
					<td width="45%" class=tablebody2>
						<table cellspacing=1 cellpadding=0 border=0 align=center width=100%>
							<tr>
								<th height=25 align="center" colspan=2>十大股东持股比例图</th>
							</tr> 
							<%if UserNum<=0 then
								response.Write "<tr valign=middle align=center height=25><td class=tablebody1 height=" & VHeight+20 & ">暂时没有人购买该股票</td></tr>"
							else
								for i=0 to Tnum
									%><tr height="<%=(VHeight+20)/UserNum%>">
										<td class=tablebody2 align="center" valign="middle" width="25%" nowrap><%=GPScale(i,0)%></td>
										<td class=tablebody1 width="75%" nowrap><img src="images/img_gupiao/baru.gif" height=10 width="<%=GPScale(i,3)+5%>" alt="持股数：<%=GPScale(i,1)%>股 持股百分比：<%=formatnumber(GPScale(i,2)*100,2,true)%>%"></td>
									</tr><%
								next
							end if%>
						</table>
					</td>
				</tr>
			</table>				 
			<br>	
			<table class=tableborder1 cellspacing=1 cellpadding=3 align=center border=0 width="<%=Forum_body(12)%>">
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
				<%set rsu=gp_conn.execute("select * from [DaHu] where sid="&sid&" and uid>0 order by ChiGuShu desc") 
				if rsu.eof and rsu.bof then 
					response.Write "<tr><td class=tablebody1 colspan=9>暂时没有人购买这个股票</td></tr>"
				else
					dim GupiaoValue,GupiaoCost
					do while not rsu.eof 
						GupiaoValue=rs("DangQianJiaGe")*rsu("ChiGuShu")		'持股现值
						GupiaoCost =rsu("PingJunJiaGe")*rsu("ChiGuShu")		'持股成本%>
					  <TR align=center valign="middle" > 
						  <td class=tablebody1 height=25><a href='z_gp_Dispu.asp?Uid=<%=rsu("uid")%>' <%if membername=rsu("ZhangHao") then%> class="cblue" <%end if%>><%=rsu("ZhangHao")%></a></TD>
						  <td class=tablebody1><%=rsu("ChiGuShu")%></TD>
						  <td class=tablebody1 align=right><%=formatnumber(rsu("ChiGuShu")*100/TotalNum,2,true)%> %&nbsp;</TD>
						  <td class=tablebody1 align=right><%=formatnumber(rs("DangQianJiaGe"),2,true)%>&nbsp;</TD>
						  <td class=tablebody1 align=right><%=formatnumber(rsu("PingJunJiaGe"),2,true)%>&nbsp;</TD>
						  <td class=tablebody1 align=right><%=formatnumber(GupiaoValue,2,true)%>&nbsp;</TD>  
						  <td class=tablebody1 align=right><%=formatnumber(GupiaoCost,2,true)%>&nbsp;</TD>
						  <td class=tablebody1 align=right><%=formatnumber(GupiaoValue-GupiaoCost,2,true)%>&nbsp;</TD>   
						  <td class=tablebody1 align=right><%=formatnumber((GupiaoValue-GupiaoCost)*100/GupiaoCost,2,true)%> %&nbsp;</TD>  
					  </TR>
						<%rsu.movenext
					loop
				end if%>
				<tr>
					<td class=tablebody2 colspan=9 align=center height=25><a href="z_gp_Exchange.asp?sid=<%=rs("sid")%>" class="cblue" title="进入交易"><b>[交易]</b></a>&nbsp;&nbsp;<a href="z_gp_Gupiao.asp"><b>[返回股市]</b></a></td>
				</tr>
			</table>		
		<%end if
		rs.close
	end if
end if
call activeonline()
call footer()

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
	set rst=gp_conn.execute("select top "&num&" ZhangHao,ChiGuShu from [DaHu] where sid="&sid&" and uid>0 order by ChiGuShu desc") 
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