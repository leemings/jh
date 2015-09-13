<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html><head>
<META http-equiv=Content-Type content=text/html; charset=gb2312>
<META HTTP-EQUIV="luxiaoqing" CONTENT="no-cache">
<meta http-equiv="refresh" content="<%=Gupiao_Setting(1)%>">
<title><%=Gupiao_Setting(5)%>-<%=membername%></title>
<!--#include file="css.asp"-->
</head><body bgcolor="#ffffff" text="#000000" style="FONT-SIZE: 9pt" topmargin=5 leftmargin=0 oncontextmenu=self.event.returnValue=false onselectstart=self.event.returnValue=false>
<!--#include file="chufa.asp"--> 
<!--
'=========================================================
' File: stock.asp
' Version:1.0
' Date: 2003-1-20
' Script Written by Nev
'=========================================================
-->
<table width="98%" border=0 cellspacing=0 cellpadding=0 align=center bgcolor="#0066CC">
	<TR>
		<TD BACKGROUND="Images/topbg.gif" height=9 colspan=3>
	</TD>
	</TR>
	<tr>
		<td valign=center align=middle height=23 background="Images/Header.gif"><font class="big"><b><%=Gupiao_Setting(5)%>大厅</b></font></td>
	</tr>
</table>
<table border=0><tr><td height="3"></td></tr></table>
<table width="98%" border=0 cellspacing=0 cellpadding=0 align=center bgcolor="#0066CC">
<tr bgcolor="#DCE2F1"><td width="200">
	<table cellspacing=1 cellpadding=0 width="100%" border=0 bgcolor="0066CC">
		<tr bgcolor="#E4E8EF"> 
			<td height=25 background='images/lan.gif' align=center><font color="#FFFFFF"><%=Custom_Setting(0)%></font></td>
		</tr>		
		<tr valign=top bgcolor="#0"> 
			<td align=center width="100%"> 
				<table cellspacing=1 cellpadding=0 border=0 width="198" height="54">
					<tr><td></td><td align=center><%=Custom_Setting(1)%></td></tr>
				</table>			
			</td>
		</tr>
	</table>
</td>
<td width="5"></td>
<td width="*">	
	<table cellspacing=1 cellpadding=0 width="100%" border=0 bgcolor="#0066CC" align=center>
		<tr>
			<td height="25" background='images/lan.gif' align="center">&nbsp;<font color="#00CC00">〓 <a href="announcements.asp"><font color="#FFFFFF">股市公告板</font></a> 〓</font></td>  
		</tr>
		<tr>
			<td height="54" bgcolor="#E4E8EF" align="center" id="content" style="FILTER: revealTrans(Transition=12, Duration=2);">
<SCRIPT LANGUAGE='JavaScript' SRC='inc/ad.js' TYPE='text/javascript'></script>
<SCRIPT LANGUAGE='JavaScript' TYPE='text/javascript'>
messages = new Array()
mescolor = new Array()
<%
		set rs=conn.execute("select top 5 * from StockNews order by id desc")
		if rs.bof and rs.eof then
%>			
				messages[0]="欢迎光临股票交易所"
				mescolor[0]="blue"
				messages[1]="证监会提醒：股市风险莫测，谨慎入市！"
				mescolor[1]="red"
<%
		else
			i=0
			do while not rs.eof 
%>			
				messages[<%=i%>]="<% response.Write rs("Title")&" <font color=gray>("&rs("AddTime")&")</font><br>"%>"
				mescolor[<%=i%>]="<% response.Write rs("color")%>"
<%
				i=i+1
				rs.movenext
			loop
%>			
			messages[<%=i%>]="证监会提醒：股市风险莫测，谨慎入市！"
			mescolor[<%=i%>]="blue" 
<%
		end if
		rs.close		
%>
dotransition()
</SCRIPT>			
			</td>
		</tr>
	</table>
</td>
<td width="5"></td>
<td width="200">
	<table cellspacing=1 cellpadding=0 width="100%" border=0 bgcolor="0066CC">
		<tr bgcolor="#E4E8EF"> 
			
          <td height=25 background='images/lan.gif' align=center><font color="#FFFFFF">电子挂历</font></td>
		</tr>		
		<tr valign=top bgcolor="#000000"> 
			<td align=center width="100%"> 
				<table cellspacing=1 cellpadding=0 border=0 width="198" height="54">
					<tr>
						<td><img src="images/clock/cb.gif" name="a"><img src="images/clock/cb.gif" name="b"><img src="images/clock/colon.gif" name="c"><img src="images/clock/cb.gif" name="d"><img src="images/clock/cb.gif" name="e"><img src="images/clock/colon.gif" name="f"><img src="images/clock/cb.gif" name="g"><img src="images/clock/cb.gif" name="h"><img src="images/clock/cam.gif" name="j"></td>
						<td align="right"><font color="#CCCCCC"><%=WeekdayName(Weekday(Date))%></font>&nbsp;</td>
					</tr>
					<tr>
                <td colspan="2" align="right"><font color="#CCCCCC"><%=DatePart("yyyy", Date) & " 年 " & DatePart("m", Date) & " 月 " & DatePart("d", Date) & " 日"%></font>&nbsp;</td>
              </tr>
				</table>			
			</td>
		</tr>
	</table>
</td></tr>
</table>
<table border=0><tr><td height="3"></td></tr></table>
<table border="2" width="98%" bgcolor="#666666" cellspacing="1" cellpadding="3" bordercolorlight="#000000" bordercolordark="#6E6E6E"  align="center">
  <tr bgcolor="#666666"> 
    <td height="25" align="center" background="Images/lan.gif"><font color=#ffffff>代号</font></td>
	<td align="center" background="Images/lan.gif"><font color=#ffffff>股票名称</font></td>
	<td align="center" background="Images/lan.gif"><font color=#ffffff>经营者</font></td>
	<td align="center" background="Images/lan.gif"><font color=#ffffff>交易量</font></td>
 	<td align="center" background="Images/lan.gif"><font color=#ffffff>昨收盘价</font></td>
	<td align="center" background="Images/lan.gif"><font color=#ffffff>成交价</font></td>
	<td align="center" background="Images/lan.gif"><font color=#ffffff>买入笔数</font></td>
	<td align="center" background="Images/lan.gif"><font color=#ffffff>卖出笔数</font></td>
	<td align="center" background="Images/lan.gif"><font color=#ffffff>成交量</font></td>
	<td align="center" background="Images/lan.gif"><font color=#ffffff>涨跌幅</font></td>
  </tr>
<%  
	dim CurrentPrice,ZangDieRange,sstop	'股票当前价格，股票涨跌幅度，股票状态(停牌，涨停，跌停等)
	Dim ColExp 		'综合指数
	Dim color
	stats="股市大厅"
	call GPOnline()	 	'在线股民        
	sql= "select * from 股票 order by sid"         
	set rs=conn.execute(sql)         
	DO while not RS.EOF  
		if rs("当前价格")<=1 then           
			color="#ffffff"           
		elseif rs("开盘价格")<rs("当前价格") then           
			color="#FF0000"           
		elseif rs("开盘价格")>rs("当前价格") then          
			color="#00FF00"           
		elseif rs("开盘价格")=rs("当前价格") then             
			color="#FFCC00"          
		end if          
		ZangDieRange=csng((rs("当前价格")-rs("开盘价格"))/rs("开盘价格"))
		CurrentPrice=formatnumber(rs("当前价格"),2,true)
		if CurrentPrice<=Trade_Setting(11)+0 then
			sstop=4
			ZangDieRange="<b>停牌</b>"
		elseif ZangDieRange>=Trade_Setting(9)+0 then          
			sstop=1
			ZangDieRange="<b>涨停</b>"         
		elseif ZangDieRange<=-Trade_Setting(10)+0 then 
			sstop=2 
			ZangDieRange="<b>跌停</b>" 
		elseif ZangDieRange=0 then
			sstop=3
			ZangDieRange="0.00 %"
		else       
			sstop=0
		end if   
%>
  <tr bgcolor="#3E3E3E" valign="bottom"> 
    <td height="20" width="6%" align="center"><a href="exchange.asp?sid=<%=rs("sid")%>"><font color=#FFCC00><%=rs("sid")%></font></a></td>
    <td height="20" width="14%" align="center"><a href="DispCompare.asp?sid=<%=rs("sid")%>"><font color="<%=color%>"><%=rs("企业")%></font></a></td>
    <td height="20" align="center"><%=iif(rs("经营者")="","<font color=#FFffff>-</font>","<a href='dispu.asp?Uid="&rs("uid")&"'><font color=#FFCC00>"&rs("经营者")&"</font></a>")%></td>
	<td height="20" align="center"><font color=#FFCC00><%=rs("交易量")%></font></td>
    <td height="20" align="center"><font color=#FFCC00><%=formatcurrency(rs("开盘价格"),2,true)%></font></td>
    <td height="20" align="center"><font color="<%=color%>"><%=formatcurrency(rs("当前价格"),2,true)%></font></td>
    <td height="20" align="center"><%=iif(rs("买入笔数")=0,"<font color=#ffffff>0</font>","<font color=#FFCC00>"&rs("买入笔数")&"</font>")%></td>
	<td height="20" align="center"><%=iif(rs("卖出笔数")=0,"<font color=#ffffff>0</font>","<font color=#FFCC00>"&rs("卖出笔数")&"</font>")%></td>
	<td height="20" align="right"><%=iif(rs("成交量")=0,"<font color=#ffffff>0&nbsp;</font>","<font color=#FFCC00>"&rs("成交量")&"&nbsp;</font>")%></td>
    <td height="20" align="center"><font color="<%=color%>"> 
<% 
		if sstop<>0 then
			response.write ""&ZangDieRange&""
		else
			response.write ""&formatnumber(ZangDieRange*100,2,true)&" %"
			ColExp=ColExp+ZangDieRange*100		'计算综合指数
		end if 
%>
      </font></td>
  </tr>
  <%        
		rs.MoveNext         
	Loop 
	rs.close
	set rs=Nothing
%>
</table>
<%
'==============个人股票管理=============
	if HaveAccount then
%>
	<br>
	<TABLE cellSpacing=1 cellPadding=2 width="98%" align=center bgColor="#666666" border=0>
		<TR bgcolor="#666666"> 
			<TD colSpan=10 height=23 valign="middle" background="Images/header.gif">&nbsp;<font color="#0066FF">〓 <font color="#FF0000">个股管理</font> 〓</font>&nbsp;&nbsp;欢迎您，<b><%=membername%></b>，您的持股情况如下：</TD>
		</TR>
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
	Dim rst,AgvInPrice		'某股票当前价格 ,某股票平均买入价格   
	Dim MyGupiaoNum,MyGupiaoValue,MyGupiaoCost 		'个人持股总数，个人股票市值,个人股票成本             
	Dim YKStr,TotalGranRate  '盈亏情况,总盈亏率
	MyGupiaoNum=0 :	MyGupiaoValue=0 :	MyGupiaoCost=0
	sql="select d.*,g.当前价格 from [大户] d inner join [股票] g on d.sid=g.sid where d.uid="&MyUserID&" order by d.[持股数] desc"  
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		RESPONSE.Write "<tr><td colspan=10 bgcolor='#8d9aca'> 　你现在没有股票</tr>"
	else
		do while not rs.EOF   
			CurrentPrice=rs("当前价格")
			AgvInPrice=rs("平均价格")
			MyGupiaoNum=MyGupiaoNum+rs("持股数")
			MyGupiaoValue=MyGupiaoValue+rs("持股数")*CurrentPrice
			MyGupiaoCost=MyGupiaoCost+rs("持股数")*AgvInPrice	
%>
		<TR align="center" bgColor="#8d9aca"> 
			<TD><a href="exchange.asp?sid=<%=rs("sid")%>"><font color="#FFFF00"><%=rs("sid")%></font></a></TD>
			<TD><a href="exchange.asp?sid=<%=rs("sid")%>"><font color="#FFFF00"><%=rs("企业")%></font></a></TD>
			<TD><font color="#FFFF00"><%=rs("持股数")%></font></TD>
			<TD><font color="#FFFF00">￥<%=formatnumber(AgvInPrice*rs("持股数"),2,True)%></font></TD>
			<TD><font color="#FFFF00">￥<%=formatnumber(CurrentPrice,2,True)%></font></TD>
			<TD><font color="#FFFF00">￥<%=formatnumber(AgvInPrice,2,True)%></font></TD>
			<TD><font color="#FFFF00">￥<%=formatnumber(CurrentPrice*rs("持股数"),2,True)%></font></TD>
			<TD>
<%
			dim bj,TotalGain,GranRate		'单股盈亏,单个股票总盈亏,单个股票总盈亏率
			bj=CurrentPrice-AgvInPrice
			bj=formatnumber(bj,2,True)
			if bj>0 then
				response.write "<font color=#FF0000>↑"&bj&" 元</font>"
			elseif bj<0 then
				response.write "<font color=#00FF00>↓"&formatnumber(abs(bj),2,True)&" 元</font>"
			elseif bj=0 then
				response.write "<font color=#FFFFff>→ 0.000 元</font>"
			end if
			response.Write "</TD><TD>"
			
			TotalGain=bj*rs("持股数")
			GranRate=TotalGain*100/MyGupiaoCost
			if TotalGain>0 then 
				response.Write "<font color=#FF0000>↑"&formatnumber(TotalGain,2,True)&" 元</font></TD>" 
				response.Write "<TD><font color=#FF0000>↑"&formatnumber(GranRate,2,True)&" %</font></TD></tr>"
			elseif TotalGain<0 then
				response.Write "<font color=#00FF00>↓"&formatnumber(-TotalGain,2,True)&" 元</font></TD>"
				response.Write "<TD><font color=#00FF00>↓"&formatnumber(-GranRate,2,True)&" %</font></TD></tr>" 
			else
				response.Write "<font color=#FFFFff>→ 0.00 元</font></TD>" 
				response.Write "<TD><font color=#FFFFff>→ 0.00 %</font></TD></tr>"	
			end if	
			rs.MoveNext          
		Loop
		rs.close
	end if
	if MyGupiaoValue-MyGupiaoCost>0 then
		YKStr="<font color=red>+ "&formatnumber(MyGupiaoValue-MyGupiaoCost,2,-1)
	elseif MyGupiaoValue-MyGupiaoCost<0 then
		YKStr="<font color=#0066FF>- "&formatnumber(MyGupiaoCost-MyGupiaoValue,2,-1)
	else
		YKStr="<font color=#9900FF>0.00"
	end if
	YKStr=YKStr&" 元</font>" 
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
			<TD background="Images/title.gif"><%=YKStr%></TD> 
			<TD background="Images/title.gif"><font color="#3366CC"><%if TotalGranRate>0 then%>+<%end if%><%=formatnumber(TotalGranRate,2,-1)%> %</font></TD>
		</TR>
	</TABLE>		
<%		
end if	
%> 
<br>
<table cellspacing=1 cellpadding=0 width="98%" align=center border=0>
	<tr>
	<td valign=top width="29%"> 		<!--第一列-->
	
		<table cellspacing=1 cellpadding=0 width="100%" border=0 bgcolor="0066CC">
			<tr bgcolor="#E4E8EF"> 
				<td height=25 background='images/lan.gif' align=center><font color="#FFFFFF">今日股市行情</font></td>
			</tr>		
        	<tr valign=top bgcolor="#E4E8EF"> 
          		<td align=center width="100%"> 
					<table cellspacing=1 cellpadding=0 border=0 width="200">
						<tr><td colspan="2" height="3"></td></tr>
						<tr><td>成交总额：</td><td align=right><%=ChengJiaoMoney%> 元</td></tr>
						<tr><td>成交笔数：</td><td align=right><%=ChengJiaoNum%> 笔</td></tr>
						<tr><td>综合指数：</td><td align=right><%if ColExp>0 then%><font color=#FF0000>↑<%elseif ColExp<0 then%><font color="#0066FF">↓<%else%><font color="#0066FF">→ <%end if%><%=formatnumber(abs(ColExp),2,true)%>%</font></font></font></td></tr>
						<tr><td colspan="2" height="3"></td></tr>
					</table>			
				</td>
			</tr>
		</table>
        <br>
		<table cellspacing=1 cellpadding=0 border=0 bgcolor="#0066CC" width="100%">
			<tr> 
				<td colspan="2"  height=25 background='images/lan.gif' align=center> <font color="#FFFFFF">个人情况总揽</font></td>
			</tr>		
<%
		if HaveAccount then			
%>          
			<tr bgcolor="#E4E8EF"> 
				<td align=center width="100%"> 
					<table cellspacing=1 cellpadding=0 border=0 width="200"> 
						<tr><td colspan="2" height="3"></td></tr>
						<tr>
							<td>帐户资金：</td> 
							<td align=right><%=formatnumber(MyCash,2,-1)%> 元</td>
						</tr>
						<tr>
							<td>持股总数：</td>  
							<td align=right><%=formatnumber(MyGupiaoNum,0,-1)%> 股</td>
						</tr>						
						<tr>
							<td>持股市值：</td>  
							<td align=right><%=formatnumber(MyGupiaoValue,2,-1)%> 元</td>
						</tr>
						<tr>
							<td>股票成本：</td>
							<td align=right><%=formatnumber(MyGupiaoCost,2,-1)%> 元</td>
						</tr>
						<tr>
							<td>盈亏情况：</td>
							<td align=right><%=YKStr%></td>
						</tr>												
						<tr>
							<td>资产合计：</td>  
							<td align=right><%=formatnumber(MyGupiaoValue+MyCash,2,-1)%> 元</td>  
						</tr>
						<tr> 
							<td>今日买卖：</td> 
							<td align=right><%=MyBiShu%> 笔</td> 
						</tr>
						<tr><td colspan="2" height="3"></td></tr>	 
					</table>
				</td>
			</tr>			
			<tr>
				<td align=center background="images/title.gif" height=21 bgcolor="#E4E8EF"><a href="Mymanage.asp" class="cblue">账户管理</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="guide/guide1.asp" class="cred">炒股指南</a></td>
			</tr>
					
		<% else %>
			<form name="form1" method="post" action="Mymanage.asp?action=new">
			<tr>
				<td width=250 bgcolor="#ffffff"> 
					<input type=hidden name=username value="<%=session("aqjh_name")%>">
					<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color="#9933CC">你在股票市场里还没有帐户，要买卖股票请先开户。</font><br><br>
				</td> 
			</tr> 
			<tr><td align=center bgcolor="#E4E8EF"><input type="submit" name="submit" value="开户" ></td></tr>
			</form>
		<% end if %>	
		</table>	
	</td>
	
	<td width=5></td>		<!--第一列和第二列之间的分隔列-->
	
    <td width="40%" align=right valign=top> 					<!-- 第二列--市实时新闻  -->
	
		<table cellspacing=1 cellpadding=0 border=0 width="100%" bgcolor="0066CC">
			<tr> 
				<td height=25 background='images/lan.gif' align=center><font color="#FFFFFF">股市实时新闻</font></td>
			</tr>
			<tr bgcolor="#444444"> 
				<td align="center">
					<MARQUEE onmouseover=this.stop() onmouseout=this.start() scrollAmount=1 scrollDelay=10 direction=up width="98%" height=206 align="left" border="0"> 
<%
					set rs=server.createobject("adodb.recordset")
					sql="select content,Addtime from RndEvent order by ID desc"
					rs.open sql,conn
					do while Not RS.Eof
						response.Write ""&rs(0)&"<br><font color=""#FFCC99""><div align=""right"">"&rs(1)&"</div><br></font>"
						RS.MoveNext
					Loop
					rs.close
					
					
					
%>					
					</marquee>
				</td>
			</tr>
			<tr>
				<td align=center background="images/footer.gif" height=25 bgcolor="#E4E8EF"><a href="news.asp" class="cblue" title="查看所有新闻">更多新闻</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="PStockApply.asp" class="cblue">个股上市</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="javascript:window.close()" class="cblue">离开股票交易中心</a></td>
			</tr>
		</table>

	</td>
		  
	<td width=5></td>		<!--第二列和第三列之间的分隔列-->
		  
	<td width="30%" align="right" valign="top"> 	<!--第三列-->
	
		<table cellspacing=1 cellpadding=0 border=0 width="100%" bgcolor="0066CC">
			<tr align=center> 
				<td height=25 background='images/lan.gif'> <font color="#FFFFFF">股市总资产排行榜</font> </td>
			</tr>
			<tr bgcolor="#E4E8EF"> 
				<td> 
					<table height=22 cellspacing=0 cellpadding=1 width="100%" border=0>	
<%
					set rs=conn.execute("select top 12 帐号,总资金,id from [客户] order by [总资金] desc")
					if RS.Eof then 
						response.Write "<td colspan=4> 还未有人买入股票 </td>"
					else 
						i=0
						response.Write "<tr><td colspan=""4"" height=""10""></td></tr>"
						do while Not RS.Eof
							i=i+1
							response.Write "<tr><td align=""center"" width=""40""><font color=red>"&i&"</font></td><td width=""5""></td>"	
							response.Write "<td><a href='dispu.asp?uid="&rs(2)&"&username="&rs(0)&"'>"&rs(0)&"</a></td>"&_
										"<td align=right>"
							if rs(1)<=0 then 
								response.Write "0 元"
							else
								response.Write formatnumber(rs(1),0,-1)&" 元&nbsp;"
							end if
							response.Write "</td></tr>"
							RS.MoveNext
							if i>=12 then exit do
						Loop
						rs.close 
					end if	
%>
					<tr><td colspan="4" height="8"></td></tr>
				</table>
			</td>
			</tr>
			<tr> 
				<td colspan=4 align=center background="images/title.gif" height=21 bgcolor="#E4E8EF" valign="middle"><a href="TopUser.asp?action=Reflash" class="cblue">刷新排名</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="TopUser.asp" class="cblue">详细排名</a></td>
			</tr>			
		</table>
	</td>
	</tr>
</table>

<% if master then %>
<br>
<table border="0" width="98%" cellspacing="1" cellpadding="3" align="center" bgcolor="#0066CC">
	<tr> 
		<td height=23 background='images/header.gif'>&nbsp;<font color="#FF0000">管理选项</font></td>
	</tr>
	<tr>
		<td height="26" valign="middle" bgcolor="#E4E8EF"><font color="#0066FF">〓 <a href="Admin_Price.asp?action=UpPrice" title="提升所有股票的现价">全红</a> ｜ <a href="Admin_Price.asp?action=DownPrice" title="降低所有股票的现价">全绿</a> ｜ <a href="Admin_Setting.asp">基本设置</a> ｜ <a href="Admin_Gupiao.asp">股票管理</a> ｜ <a href="Admin_Gupiao.asp?action=newgupiao">新股上市</a> ｜ <a href="Admin_User.asp">客户管理</a> ｜ <a href="Admin_Data.asp">数据库管理</a> 〓</font></td>      
	</tr>
</table> 
<%end if%>  
<br>
<table cellspacing=1 cellpadding=0 width="98%" border=0 bgcolor="#0066CC" align=center>
	<tr><td height="23" background='images/lan.gif'><font color="#FFFFFF">&nbsp;在线股民</font></td></tr>
	<%call onlineuser(1,1)%>
</table> 
<br>
<table cellspacing=1 cellpadding=2 width="98%" border=0 bgcolor="#0066CC" align=center>
	<tr>  
		<td  height=32 background="images/footer.gif" valign=top>
			<table border=0 width="100%" cellspacing=0 cellpadding=5 bordercolorlight="#ffffff" bordercolordark="#ffffff">
				<tr valign="bottom"><td id="clock" width="185" align="left">欢迎您</td><td width="*" id="CopyRight" align="center">制作维护：快乐江湖总站</td><td width="170" align="right">页面执行时间：<%=FormatNumber((timer()-startime)*1000,3)%>毫秒</td></tr>
			</table>	
		</td>
	</tr>
</table> 
<% 
CloseDatabase		'关闭股票数据库

%>
</body>
</html>
<!--#include file="inc/js.asp"-->