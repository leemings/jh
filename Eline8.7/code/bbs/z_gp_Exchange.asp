<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_gp_conn.asp"-->
<!--#include file="z_gp_Const.asp"-->
<%
stats="股票交易"
call nav()
call head_var(0,0,GuPiao_Setting(5),"z_gp_gupiao.asp")
response.write "<table class=tableborder1 cellspacing=1 cellpadding=3 align=center border=0 width=" & Forum_body(12) & ">"
if not HaveAccount then
	errmess="<li>您还没有股票帐户，不能炒股，请先开户"
	call endinfo(1)
else
	dim sid,mess,BuyTime,SaleTime
	if Request.Form("sid")<>"" then
		sid=trim(Request.Form("sid"))
	else	
		sid=trim(Request.QueryString("sid"))
	end if
	if sid="" or (not isnumeric(sid)) then
		errmess="<li>非法炒股，请从有效连接进入！"
		call endinfo(1)
	else
		sql="select JiaoYiLiang,DangQianJiaGe,QiYe,ZhuangTai,KaiPanJiaGe from [GuPiao] where sid="&sid
		set rs=gp_conn.execute(sql)
		if rs.eof and rs.bof then
			rs.close
			errmess="<li>没有这个股票"
			call endinfo(1)
		elseif rs(3)="封" then
			rs.close
			errmess="<li>该股票已经封盘，不能进行交易"
			call endinfo(2)
		else
			dim NowPrice,NowNum,GupiaoName,OpenPrice,RaisingLimit,FallLimit
			NowNum=rs(0)		'交易量
			NowPrice=rs(1)		'当前价格
			OpenPrice=rs(4)		'开盘价格 
			GupiaoName=rs(2)	'股票名称 
			rs.close
			RaisingLimit=False	'股票是否处于涨停状态 
			FallLimit=False		'股票是否处于跌停状态 
			if NowPrice>=9999 OR (NowPrice-OpenPrice)/OpenPrice>=Trade_Setting(9)+0 then
				RaisingLimit=True	'涨停板 能卖出不能买入
			elseif NowPrice<=Trade_Setting(11)+0 OR (OpenPrice-NowPrice)/OpenPrice>=Trade_Setting(10)+0 then
				FallLimit=True		'跌停板 能买入不能卖出
			end if
			select case request("action")
				case "buy"
					call Buy()
				case "sale"
					call Sale()
				case else					
					call main()
			end select	
		end if
	end if
end if
set rs=nothing
response.write "</table>"
call activeonline()
call footer()
'======================================
sub main()
	dim sushare,BuyMax
	sql="select ChiGuShu from [DaHu] where uid="&MyUserID&" and sid="&sid
	set rs=gp_conn.execute(sql)
	if rs.eof then 
		sushare=0
	else
		sushare=rs(0)
	end if	
	'计算最多可以购买该股票多少股
	if NowNum<(int(MyCash*(1-Trade_Setting(5)/100)/nowprice)) and MyCash>=NowNum*nowprice*(1+Trade_Setting(5)/100) then   
		BuyMax=NowNum
	else
		BuyMax=int(MyCash*(1-Trade_Setting(5)/100)/nowprice) 
	end if%>
	<tr>
		<th valign=center align=middle height=23>第 <font color="<%=forum_body(8)%>"><%=sid%></font> 号股票<%=GupiaoName%></th>
	</tr>
	<tr>
		<td class=tablebody1 valign=center>
			<form name="form1" method="post" action="?action=buy" >
			<input type="hidden" name="sid" value="<%=sid%>"><br>
		  <div align="center"> 
      	<p><font color="#006633">股票数量：<font color="#ff0000"><%=NowNum%></font> 股当前价格：<font color="#ff0000"><%=formatnumber(nowprice,3)%></font> 元</font></p>
  		</div>
  		<div align="center"> 
      	<p><font color="#3366CC">您的股票帐户里现有资金： <font color="#9900CC"><%=formatnumber(MyCash,0)%> </font>元</font></p>
  		</div>
  		<div align="center"> 
	      <p><font color="#006633">购买数量：</font> 
        <input name="BuyNum" size="20" value="<%=BuyMax%>" onFocus="this.select()">
        <input type="submit" value="购买" name="B1">
        <input  type="reset" value="重填" name="B2">
        <br><br><font color="#990000">扣除 <%=Trade_Setting(5)%>% 的手续费后，您最多可以购买 <font color="#FF0000"><%=BuyMax%> </font>股</font></p>
  		</div>
			</form>
			<form name="form2" method="post" action="?action=sale">
			<input type="hidden" name="sid" value="<%=sid%>">
			<div align="center">
				<font color="#006633">卖出数量：</font>
				<input name="SaleNum" size="20" value="<%=sushare%>" onFocus="this.select()">
				<input type="submit" value="卖出" name="B3">
				<input type="reset" value="重填" name="B2">
			</div>
			<br>
			<div align="center">
				<font color="#990000">您已经拥有这种股票：<font color="#006633"><%=sushare%></font> 股</font>
			</div>
			</form>
			<div ALIGN=CENTER>
				<FONT color="#3366CC">证监会提醒：股市风险莫测，谨慎入市！</FONT>
			</div>
			<br>
		</td>
	</tr>
	<TR>
		<Td class=tablebody2 height=25 align="center"><A href="z_gp_Gupiao.asp" class="cblue"><b>[股市大厅]</b></A>&nbsp;&nbsp;<A href="javascript:history.go(-1)" class="cblue"><b>[返回上一页]</b></a></Td>
	</TR>
<%end sub
'-------------------买入股票--------------------------
sub Buy() 
	if Trade_Setting(0)=0 then 
		errmess="<li>股市交易已经暂停，不能进行股票买卖操作"
		call endinfo(1) 
		exit sub	
	end if 
	if RaisingLimit and cint(Trade_Setting(16))<>3 then
		if cint(Trade_Setting(16))<>0 then
			errmess="<li>本股市已经限制对于涨停的股票是不能进行买入操作"
			call endinfo(1) 
			exit sub
		end if
	end if
	if FallLimit and cint(Trade_Setting(17))<>3 then
		if cint(Trade_Setting(17))<>0 then
			errmess="<li>本股市已经限制对于跌停的股票是不能进行买入操作"
			call endinfo(1) 
			exit sub
		end if
	end if
	if not master and clng(Trade_Setting(4))>0 and MyBiShu>=clng(Trade_Setting(4)) then
		errmess="<li>本股市已经限制一天最多可以做 "&Trade_Setting(4)&" 笔交易，您已经达到了这个上限，请改天再进行交易"
		call endinfo(1) 
		exit sub	
	end if
	Dim BuyNum	'买入数量
	if Request.Form("BuyNum")="" or (not isnumeric(Request.Form("BuyNum"))) then
		errmess="<li>您要买多少股票啊？"
		call endinfo(2)
		exit sub
	end if
	BuyNum=int(Request.Form("BuyNum"))
	if BuyNum<=0 then
		errmess="<li>交易量不能小于等于 0 股！"
		call endinfo(2)
		exit sub	
	elseif not master and BuyNum<clng(Trade_Setting(2)) then
		errmess="<li>本股市已经限制每次最少交易量为 "&Trade_Setting(2)&" 股！"
		call endinfo(2)
		exit sub	 
	elseif not master and BuyNum>clng(Trade_Setting(3)) then
		errmess="<li>本股市已经限制每次最大交易量为 "&Trade_Setting(3)&" 股！"
		call endinfo(2)
		exit sub
	end if

	dim DelMoney,Poundage	'买股票需要的钱,手续费
	DelMoney=BuyNum*NowPrice
	Poundage=DelMoney*Trade_Setting(5)/100		'其中 Trade_Setting(5) 是手续费
	
	if MyCash<DelMoney-Poundage then	'2 
		errmess="<li>您的资金不足以购买指定的股票，您可以减少购买数量或者购买其它股票"
		call endinfo(1)
	elseif int(NowNum)<int(BuyNum) then
		errmess="<li>没有足够的股票数量，您可以减少购买数量或者购买其它股票"
		call endinfo(2)
	else
		Dim Rnum,NewPrice,Zd,Rate,TotalStockNum		'随机数,股票价格变化量,涨跌幅
		set rs=gp_conn.execute("select ChiGuShu,PingJunJiaGe,BuyTime from [DaHu] where sid="&sid&" and uid="&MyUserID&"")
		
		TotalStockNum=Max((gp_conn.execute("select ZongGuFen from [GuPiao] where sid="&sid)(0))/30,1)	'总股份
			
		if rs.eof and rs.bof then	'1		'如果没有买过这个股票
			Randomize
			Rate=Rnd
			Rnum=Rnd
			if Rnum<=0.4 then Rnum=Rnum+0.5
			if Rate<Trade_Setting(12)+0 then	'买入时价格上涨
				NewPrice=NowPrice*BuyNum*Rnum/TotalStockNum 		'股票价格变化量
				if NowPrice+NewPrice>OpenPrice*(1+Trade_Setting(9)) then
					NewPrice=OpenPrice*(1+Trade_Setting(9))-NowPrice
				elseif NowPrice+NewPrice<OpenPrice*(1-Trade_Setting(10)) then
					NewPrice=OpenPrice*(1-Trade_Setting(10))-NowPrice
				end if
				ZD=NewPrice/NowPrice
				mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>买入"&GupiaoName&" "&BuyNum&" 股，现价上涨 "&formatpercent(Zd,2,-1)&"</font>"
			elseif Rate<Trade_Setting(12)+0+Trade_Setting(13) then	'买入时价格下滑
				NewPrice=-NowPrice*BuyNum*Rnum/TotalStockNum		'价格变化值 +
				if NowPrice+NewPrice>OpenPrice*(1+Trade_Setting(9)) then
					NewPrice=OpenPrice*(1+Trade_Setting(9))-NowPrice
				elseif NowPrice+NewPrice<OpenPrice*(1-Trade_Setting(10)) then
					NewPrice=OpenPrice*(1-Trade_Setting(10))-NowPrice
				end if
				ZD=NewPrice/NowPrice 
				mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>买入"&GupiaoName&" "&BuyNum&" 股，现价下滑 "&formatpercent(-Zd,2,-1)&"</font>"			
			else
				NewPrice=0		'价格变化值 0 
				ZD=0
				mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>买入"&GupiaoName&" "&BuyNum&" 股，现价没有波动</font>"						
			end if
			if ChangeOptB(sid,MyUserID,membername,BuyNum)=true then mess=mess+"<font color=#33CCFF>，并取得经营权</font>"
			
			gp_conn.execute("update [GuPiao] set DangQianJiaGe=DangQianJiaGe+"&NewPrice&",ShengYuGuFen=ShengYuGuFen-"&BuyNum&",JiaoYiLiang=JiaoYiLiang-"&BuyNum&",ChengJiaoLiang=ChengJiaoLiang+"&BuyNum&",MaiRuBiShu=MaiRuBiShu+1,TodayWave=TodayWave+"&ZD&",TotalWave=TotalWave+"&ZD&" where sid="&sid)
			gp_conn.execute("insert into [DaHu] (Uid,ZhangHao,sid,MaiRuJiaGe,PingJunJiaGe,ChiGuShu,QiYe) values ("&MyUserID&",'"&membername&"','"&sid&"','"&NowPrice&"','"&NowPrice&"','"&BuyNum&"','"&GupiaoName&"')")
			gp_conn.execute("update [KeHu] set ZiJin=ZiJin-"&(DelMoney+Poundage)&",ChiGuZhongLei=ChiGuZhongLei+1,JinRiMaiRu=JinRiMaiRu+1,ZongMaiRu=ZongMaiRu+1 where  ID="&MyUserID)
			gp_conn.execute("insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )")  
			
		elseif not master and datediff("n",rs(2),now())<Trade_Setting(6) then
			errmess="<li>本股市已经限制两次买入同一股票的时间间隔为："&Trade_Setting(6)&" 分钟<br><li>您上次买入该股票的时间是 "&rs(2)&",离下次可以购买该股票还差："&(Trade_Setting(6)-datediff("n",rs(2),now()))&" 分钟"
			founderr=true
			call endinfo(1)
			exit sub
			
		else
			dim MyNum,AgvPrice	'持股数，平均价格
			MyNum=rs(0)			'已有股票数目			
			AgvPrice=rs(1)		'平均价格
			Randomize
			Rate=Rnd
			Rnum=Rnd
			if Rnum<=0.4 then Rnum=Rnum+0.5
			if Rate<Trade_Setting(12)+0 then	'买入时价格上涨
				NewPrice=NowPrice*BuyNum*Rnum/TotalStockNum		'股票价格变化量
				if NowPrice+NewPrice>OpenPrice*(1+Trade_Setting(9)) then
					NewPrice=OpenPrice*(1+Trade_Setting(9))-NowPrice
				elseif NowPrice+NewPrice<OpenPrice*(1-Trade_Setting(10)) then
					NewPrice=OpenPrice*(1-Trade_Setting(10))-NowPrice
				end if
				ZD=NewPrice/NowPrice
				mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>买入"&GupiaoName&" "&BuyNum&" 股，现价上涨 "&formatpercent(Zd,2,-1)&"</font>"
			elseif Rate<Trade_Setting(12)+Trade_Setting(13)+0 then	'买入时价格下滑
				NewPrice=-NowPrice*BuyNum*Rnum/TotalStockNum		'价格变化值
				if NowPrice+NewPrice>OpenPrice*(1+Trade_Setting(9)) then
					NewPrice=OpenPrice*(1+Trade_Setting(9))-NowPrice
				elseif NowPrice+NewPrice<OpenPrice*(1-Trade_Setting(10)) then
					NewPrice=OpenPrice*(1-Trade_Setting(10))-NowPrice
				end if
				ZD=NewPrice/NowPrice
				mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>买入"&GupiaoName&" "&BuyNum&" 股，现价下滑 "&formatpercent(-Zd,2,-1)&"</font>"			
			else
				NewPrice=0		'价格变化值 0 
				ZD=0
				mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>买入"&GupiaoName&" "&BuyNum&" 股，现价没有波动</font>"						
			end if
			if ChangeOptB(sid,MyUserID,membername,MyNum+BuyNum)=true then mess=mess+"<font color=#33CCFF>，并取得经营权</font>"

			AgvPrice=(NowPrice*BuyNum+AgvPrice*MyNum)/(BuyNum+MyNum)  	'新平均价格			
			
			gp_conn.execute("update [GuPiao] set DangQianJiaGe=DangQianJiaGe+"&NewPrice&",ShengYuGuFen=ShengYuGuFen-"&BuyNum&",JiaoYiLiang=JiaoYiLiang-"&BuyNum&",ChengJiaoLiang=ChengJiaoLiang+"&BuyNum&",MaiRuBiShu=MaiRuBiShu+1,TodayWave=TodayWave+"&ZD&",TotalWave=TotalWave+"&ZD&" where sid="&sid)
			gp_conn.execute("update [DaHu] set PingJunJiaGe="&AgvPrice&",MaiRuJiaGe="&NowPrice&",ChiGuShu=ChiGuShu+"&BuyNum&",BuyTime=now() where Uid="&MyUserID&" and sid="&sid)
			gp_conn.execute("update [KeHu] set ZiJin=ZiJin-"&(DelMoney+Poundage)&",JinRiMaiRu=JinRiMaiRu+1,ZongMaiRu=ZongMaiRu+1 where  ID="&MyUserID)
			gp_conn.execute("insert into RndEvent(content,addtime) values('"&mess&"',now())")  
		end if
		call GivePoundage(sid,Poundage)
		call CalculateFund(MyUserID)
		gp_conn.execute("update [GupiaoConfig] set TodayBuy=TodayBuy+1,TodayTotal=TodayTotal+"&DelMoney&"")
		sucmess="<li>股票购买成功！"
		call endinfo(1)
		if Gupiao_Setting(12)>0 then
			'只保留最新Gupiao_Setting(12)个事件
			gp_conn.execute("delete from RndEvent where id not in(select top "&Gupiao_Setting(12)&" id from RndEvent order by id desc)")
		end if
		
	end if	'2
end sub
'-------------------卖出股票--------------------
sub Sale()
	if Trade_Setting(0)=0 then 
		errmess="<li>股市交易已经暂停，不能进行股票买卖操作"
		call endinfo(1) 
		exit sub	
	end if
	if RaisingLimit and cint(Trade_Setting(16))<>3 then
		if cint(Trade_Setting(16))<>1 then
			errmess="<li>本股市已经限制对于涨停的股票是不能进行卖出操作"
			call endinfo(1) 
			exit sub
		end if
	end if
	if FallLimit and cint(Trade_Setting(17))<>3 then
		if cint(Trade_Setting(17))<>1 then
			errmess="<li>本股市已经限制对于跌停的股票是不能进行卖出操作"
			call endinfo(1) 
			exit sub
		end if
	end if	
	if not master and clng(Trade_Setting(4))>0 and MyBiShu>=clng(Trade_Setting(4)) then
		errmess="<li>本股市已经限制一天最多可以做 "&Trade_Setting(4)&" 笔交易，您已经达到了这个上限，请改天再进行交易"
		call endinfo(1) 
		exit sub	
	end if	
	Dim SaleNum,MyNum,Rnum,Rate,TotalStockNum	'卖出数量，交易量数量，自己还剩下数量，随机数
	Dim AddMoney,NewPrice,ZD,Rname,Poundage
	 
	if Request.Form ("SaleNum")="" or (not isnumeric(Request.Form ("SaleNum"))) then
		errmess="<li>请输入卖出股票的数量"
		call endinfo(2)
		exit sub
	elseif Request.Form ("SaleNum")<=0 then
		errmess="<li>卖出的股票数量不能为零或者负数"
		call endinfo(2)
		exit sub
	end if
	SaleNum=int(Request.Form ("SaleNum"))
		
	AddMoney=SaleNum*NowPrice
	Poundage=AddMoney*Trade_Setting(5)/100 	'手续费
		
	set rs=gp_conn.execute ("select ChiGuShu,BuyTime,SaleTime from [DaHu] where Uid="&MyUserID&" and sid="&sid&"")
	
	if rs.eof and rs.bof then 
		errmess="<li>您没有这个股票，怎么卖啊？"
		founderr=true
	elseif not master and datediff("n",rs(1),now())\60<Trade_Setting(8) then
		errmess="<li>本股市已经限制买入某股票后再卖出的时间间隔为："&Trade_Setting(8)&" 小时<br><li>您上次买入该股票的时间是 "&rs(1)&",离可以抛出该股票还差："&(Trade_Setting(8)*60-datediff("n",rs(1),now()))&" 分钟"
		founderr=true				
	elseif not master and datediff("n",rs(2),now())<Trade_Setting(7) then
		errmess="<li>本股市已经限制两次卖出同一股票的时间间隔为："&Trade_Setting(7)&" 分钟<br><li>您上次卖出该股票的时间是 "&rs(2)&" 分钟,离下次可以抛出该股票还差："&(Trade_Setting(7)-datediff("n",rs(2),now()))&" 分钟"
		founderr=true
	elseif not master and SaleNum<Trade_Setting(2)+0 and rs("ChiGuShu")>Trade_Setting(2)+0 then
		errmess="<li>本股市已经限制每次最少交易量为 "&Trade_Setting(2)&" 股<br><li>如果您的股票数不足 "&Trade_Setting(2)&" 股，您可以全部卖出"
		founderr=true
	elseif not master and SaleNum>Trade_Setting(3)+0 then
		errmess="<li>本股市已经限制每次最大交易量为 "&Trade_Setting(3)&" 支股票！"
		founderr=true
	else 			
		MyNum=rs("ChiGuShu")-SaleNum		
	end if
	rs.close
	
	if not founderr then
		set rs=gp_conn.execute("select ZongGuFen from [GuPiao] where sid="&sid)
		if rs.eof and rs.bof then
			errmess="<li>该股票已经删除，请与管理员联系"
			founderr=true		
		else
			TotalStockNum=Max(rs(0)/30,1)	'总股份
		end if
	end if
	
	if founderr then
		call endinfo(1)
		exit sub
	end if	
			
	if MyNum<0 then			'没有这么多股票
		errmess="<li>您没有这么多股票，怎么卖啊？"
		call endinfo(2)
		exit sub	
	elseif MyNum=0 then		'全部卖出 
		Randomize
		Rate=Rnd
		Rnum=Rnd
		if Rnum<=0.4 then Rnum=Rnum+0.5
		if Rate<Trade_Setting(14)+0 then	'卖出时价格下滑
			NewPrice=-NowPrice*SaleNum*Rnum/TotalStockNum		'价格变化值 -
			if NowPrice+NewPrice>OpenPrice*(1+Trade_Setting(9)) then
				NewPrice=OpenPrice*(1+Trade_Setting(9))-NowPrice
			elseif NowPrice+NewPrice<OpenPrice*(1-Trade_Setting(10)) then
				NewPrice=OpenPrice*(1-Trade_Setting(10))-NowPrice
			end if
			ZD=NewPrice/NowPrice 
			mess="<font color=#FF6666>"&membername&"</font><font color=#33CCFF>卖出"&GupiaoName&" "&SaleNum&" 股，现价下滑 "&formatpercent(-Zd,2,-1)&"</font>"
		elseif Rate<Trade_Setting(14)+Trade_Setting(15)+0 then	'卖出时价格上涨
			NewPrice=NowPrice*SaleNum*Rnum/TotalStockNum		'价格变化值 +
			if NowPrice+NewPrice>OpenPrice*(1+Trade_Setting(9)) then
				NewPrice=OpenPrice*(1+Trade_Setting(9))-NowPrice
			elseif NowPrice+NewPrice<OpenPrice*(1-Trade_Setting(10)) then
				NewPrice=OpenPrice*(1-Trade_Setting(10))-NowPrice
			end if
			ZD=NewPrice/NowPrice 
			mess="<font color=#FF6666>"&membername&"</font><font color=#33CCFF>卖出"&GupiaoName&" "&SaleNum&" 股，现价上涨 "&formatpercent(Zd,2,-1)&"</font>"			
		else
			NewPrice=0		'价格变化值 0 
			ZD=0
			mess="<font color=#FF6666>"&membername&"</font><font color=#33CCFF>卖出"&GupiaoName&" "&SaleNum&" 股，现价没有波动</font>"						
		end if			
		
		gp_conn.execute "update [GuPiao] set ChengJiaoLiang=ChengJiaoLiang+"&SaleNum&",ShengYuGuFen=ShengYuGuFen+"&SaleNum&",DangQianJiaGe=DangQianJiaGe+"&NewPrice&", JiaoYiLiang=JiaoYiLiang+"&SaleNum&",MaiChuBiShu=MaiChuBiShu+1,TodayWave=TodayWave+"&ZD&",TotalWave=TotalWave+"&ZD&" where sid="&sid
		gp_conn.execute "delete from [DaHu] where Uid="&MyUserID&" and sid="&sid
		gp_conn.execute "update [KeHu] set ZiJin=ZiJin+"&AddMoney-Poundage&",ChiGuZhongLei=ChiGuZhongLei-1,JinRiMaiChu=JinRiMaiChu+1,ZongMaiChu=ZongMaiChu+1 where ID="&MyUserID

		Rname=ChangeOptS(sid,MyUserID,membername,0)
		if Rname<>"" then mess=mess&"<font color=#33CCFF>,<font color=#FF6666>"&Rname&"</font>获得该股的经营权</font>"
		gp_conn.execute "insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )"  
	else
		Randomize
		Rate=Rnd
		Rnum=Rnd
		if Rnum<0.4 then Rnum=Rnum+0.5
		if Rate<Trade_Setting(14)+0 then	'卖出时价格下滑
			NewPrice=-NowPrice*SaleNum*Rnum/TotalStockNum		'价格变化值 -
			if NowPrice+NewPrice>OpenPrice*(1+Trade_Setting(9)) then
				NewPrice=OpenPrice*(1+Trade_Setting(9))-NowPrice
			elseif NowPrice+NewPrice<OpenPrice*(1-Trade_Setting(10)) then
				NewPrice=OpenPrice*(1-Trade_Setting(10))-NowPrice
			end if
			ZD=-NewPrice/NowPrice 
			mess="<font color=#FF6666>"&membername&"</font><font color=#33CCFF>卖出"&GupiaoName&" "&SaleNum&" 股，现价下滑 "&formatpercent(Zd,2,-1)&"</font>"
		elseif Rate<Trade_Setting(14)+Trade_Setting(15)+0 then	'卖出时价格上涨
			NewPrice=NowPrice*SaleNum*Rnum/TotalStockNum		'价格变化值 +
			if NowPrice+NewPrice>OpenPrice*(1+Trade_Setting(9)) then
				NewPrice=OpenPrice*(1+Trade_Setting(9))-NowPrice
			elseif NowPrice+NewPrice<OpenPrice*(1-Trade_Setting(10)) then
				NewPrice=OpenPrice*(1-Trade_Setting(10))-NowPrice
			end if
			ZD=NewPrice/NowPrice 
			mess="<font color=#FF6666>"&membername&"</font><font color=#33CCFF>卖出"&GupiaoName&" "&SaleNum&" 股，现价上涨 "&formatpercent(Zd,2,-1)&"</font>"			
		else
			NewPrice=0		'价格变化值 0 
			ZD=0
			mess="<font color=#FF6666>"&membername&"</font><font color=#33CCFF>卖出"&GupiaoName&" "&SaleNum&" 股，现价没有波动</font>"						
		end if		
		gp_conn.execute "update [GuPiao] set ChengJiaoLiang=ChengJiaoLiang+"&SaleNum&",ShengYuGuFen=ShengYuGuFen+"&SaleNum&",DangQianJiaGe=DangQianJiaGe+"&NewPrice&", JiaoYiLiang=JiaoYiLiang+"&SaleNum&",MaiChuBiShu=MaiChuBiShu+1,TodayWave=TodayWave+"&ZD&",TotalWave=TotalWave+"&ZD&" where sid="&sid
		gp_conn.execute "update [DaHu] set ChiGuShu="&MyNum&",SaleTime=now() where Uid="&MyUserID&" and sid="&sid
		gp_conn.execute "update [KeHu] set ZiJin=ZiJin+"&AddMoney-Poundage&",JinRiMaiChu=JinRiMaiChu+1,ZongMaiChu=ZongMaiChu+1 where ID="&MyUserID
	
		Rname=ChangeOptS(sid,MyUserID,membername,MyNum)
		if Rname<>"" then mess=mess&"<font color=#33CCFF>,<font color=#FF6666>"&Rname&"</font>获得该股的经营权</font>"		
		gp_conn.execute "insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )"  
	end if
	call GivePoundage(sid,Poundage)
	call CalculateFund(MyUserID)
	gp_conn.execute("update [GupiaoConfig] set TodaySale=TodaySale+1,TodayTotal=TodayTotal+"&AddMoney&"")
	sucmess="<li>股票卖出交易已成功提交！"
	call endinfo(1)
	if Gupiao_Setting(12)>0 then
		'只保留最新Gupiao_Setting(12)个事件
		gp_conn.execute("delete from RndEvent where id not in(select top "&Gupiao_Setting(12)&" id from RndEvent order by id desc)")
	end if
end sub

'判断是否取得买入股票的经营权
Function ChangeOptB(sid,Uid,UserName,TotalNum)
	ChangeOptB=false
	set rs=gp_conn.execute("select Uid from [GuPiao] where sid="&sid)
	if rs(0)=0 then		'如果该股票还没有人经营
		gp_conn.execute("update [GuPiao] set Uid="&Uid&",JingYingZhe='"&UserName&"' where sid="&sid)
		ChangeOptB=true
	elseif rs(0)<>Uid then	'如果该股票已经有其他人经营，则需要判断谁的股票多
		rs.close  
		set rs=gp_conn.execute("select top 1 ChiGuShu,Uid from [DaHu] where sid="&sid&" and Uid<>"&Uid&" order by ChiGuShu desc")  
		if not rs.eof then
			if TotalNum>rs(0) then 
				gp_conn.execute("update [GuPiao] set JingYingZhe='"&UserName&"',Uid="&Uid&" where sid="&sid)
				ChangeOptB=true
			end if	
		end if
		rs.close 	
	end if
end Function

'判断是否失去对卖出的股票的经营权,RemNum：卖出后剩余的持股数
Function ChangeOptS(sid,Uid,UserName,RemNum)
	ChangeOptS=""
	set rs=gp_conn.execute("select Uid from [GuPiao] where sid="&sid)
	if rs(0)=Uid then	'如果现在的经营权是卖股票的用户时才需要判断
		rs.close
		set rs=gp_conn.execute("select top 1 ChiGuShu,ZhangHao,Uid from [DaHu] where sid="&sid&" and Uid>0 and Uid<>"&Uid&" order by ChiGuShu desc")
		if not rs.eof then		'除了自己之外还有其他人卖这个股票时需要进行判断
			if RemNum<rs(0) then    
				gp_conn.execute("update [GuPiao] set JingYingZhe='"&rs(1)&"',Uid="&rs(2)&" where sid="&sid) 
				ChangeOptS=rs(1)
			end if
		elseif remnum=0 then		'如果只有卖者拥有这个股票而且本次卖出全部股票
			gp_conn.execute("update [GuPiao] set JingYingZhe=null,Uid=0 where sid="&sid)
		end if
		rs.close	
	end if	
end Function

'经营者 获得 手续费 
sub GivePoundage(id,HandleMoney) 
	set rs=gp_conn.execute("select Uid from [GuPiao] where sid="&id)
	if not rs.eof then
		gp_conn.execute("update [KeHu] set ZiJin=ZiJin+"&csng(HandleMoney)&",ZongZiJin=ZongZiJin+"&csng(HandleMoney)&" where id="&rs(0) )
	end if	
end sub

sub CalculateFund(Usid) 
	dim rst,TotalFund
	TotalFund=0	
	sql="select d.ChiGuShu,g.DangQianJiaGe from [DaHu] d inner join [GuPiao] g on d.sid=g.sid where d.uid="&Usid
	set rst=gp_conn.execute(sql)
	do while not rst.eof
		TotalFund=TotalFund+rst(0)*rst(1)
		rst.movenext
	loop
	rst.close
	gp_conn.execute("update [KeHu] set ZongZiJin=ZiJin+"&TotalFund&" where id="&Usid)
end sub
%>