<%dim Rnum
Randomize
Rnum=rnd
if Rnum<=Gupiao_Setting(10)+0 then '请不要修改本句，否则可能会出错
	dim RndID,mess,MaxID
	MaxID=gp_conn.execute("select top 1 cid from [GuPiao] order by cid desc")(0)		'MaxID 就是现有股票的数目  
	Randomize rnd
	RndID=int(MaxID*rnd)+1		'随机事件，这里求出事件对应的股票的CID
	sql= "select * from [GuPiao] where Cid="&RndID        
	set rs=gp_conn.execute(sql)
	if (rs("DangQianJiaGe")-rs("KaiPanJiaGe"))/rs("KaiPanJiaGe")>=Trade_Setting(9)+0 then	'涨停
		sql="update [GuPiao] set DangQianJiaGe=DangQianJiaGe*0.93 where Cid="&RndID&""
		gp_conn.execute sql           
		mess="<font color=#FFFF00>"&rs("QiYe")&"</font><font color=white>的股票遭大户抛售，打开涨停状态！</font>"
		sql="insert into RndEvent(content,addtime) values('"&mess&"',now())"
		gp_conn.execute sql  
	elseif (rs("DangQianJiaGe")-rs("KaiPanJiaGe"))/rs("KaiPanJiaGe")<=-Trade_Setting(10)+0 then	'跌停	
		sql="update [GuPiao] set DangQianJiaGe=DangQianJiaGe*1.08 where Cid="&RndID&""
		gp_conn.execute sql              
		mess="<font color=#FFFF00>"&rs("QiYe")&"</font><font color=white>的股票被大量购入，打开跌停状态！</font>"
		sql="insert into RndEvent(content,addtime) values('"&mess&"',now())"
		gp_conn.execute sql  
	else   
		Randomize
		if Rnd<Gupiao_Setting(11)+0 then		'股票价格上涨 事件
			'Randomize
			Rnum=formatnumber(100*rnd+1,2,true)		'1～100 之间的随机数
			if Rnum<10 then 
				RndEvent Rnum*10,1,rs("QiYe"),"资产重组，现价微涨","red",RndID,1 
			elseif Rnum<20 then
				RndEvent Rnum*10,1,rs("QiYe"),"资产重组，现价上涨","red",RndID,1
			elseif Rnum<30 then
				RndEvent Rnum*10,10,rs("QiYe"),"公司业绩提高，投资人信心加大，现价上涨","red",RndID,1
			elseif Rnum<40 then
				RndEvent Rnum*10,20,rs("QiYe"),"大规模资产重组，现价上涨","red",RndID,1
			elseif Rnum<50 then
				RndEvent Rnum*10,300,rs("QiYe"),"业绩大幅提高，注入大量资金，现价上涨","red",RndID,1
			elseif Rnum<60 then 
				RndEvent 1000-Rnum*10,40,rs("QiYe"),"业绩大幅提高，注入大量资金，现价上涨","red",RndID,1 
			elseif Rnum<70 then
				RndEvent 1000-Rnum*10,30,rs("QiYe"),"大规模资产重组，现价上涨","red",RndID,1
			elseif Rnum<80 then
				RndEvent 1000-Rnum*10,20,rs("QiYe"),"公司业绩提高，投资人信心加大，现价上涨","red",RndID,1
			elseif Rnum<90 then
				RndEvent 1000-Rnum*10,10,rs("QiYe"),"资产重组，现价上涨","red",RndID,1
			else
				RndEvent 1000-Rnum*10,1,rs("QiYe"),"资产重组，现价微涨","red",RndID,1
			end if
		else									'股票价格下滑 事件
			'Randomize
			Rnum=formatnumber(100*rnd+1,2,true)		'1～100 之间的随机数  
			if Rnum<10 then 
				RndEvent Rnum*10,1,rs("QiYe"),"放出非利好消息，现价微跌","#00FF00",RndID,-1 
			elseif Rnum<20 then
				RndEvent Rnum*10,1,rs("QiYe"),"放出非利好消息，现价下滑","#00FF00",RndID,-1
			elseif Rnum<30 then
				RndEvent Rnum*10,10,rs("QiYe"),"公司业绩不佳，投资人信心不足，现价下滑","#00FF00",RndID,-1
			elseif Rnum<40 then
				RndEvent Rnum*10,20,rs("QiYe"),"业绩开始滑坡，现价下跌","#00FF00",RndID,-1
			elseif Rnum<50 then
				RndEvent Rnum*10,300,rs("QiYe"),"业绩大幅滑坡，资金不足，现价下跌","#00FF00",RndID,-1
			elseif Rnum<60 then 
				RndEvent 1000-Rnum*10,40,rs("QiYe"),"业绩大幅滑坡，遭受大户抛售，现价下跌","#00FF00",RndID,-1 
			elseif Rnum<70 then
				RndEvent 1000-Rnum*10,30,rs("QiYe"),"业绩开始滑坡，现价下跌","#00FF00",RndID,-1
			elseif Rnum<80 then
				RndEvent 1000-Rnum*10,20,rs("QiYe"),"公司业绩不佳，投资人信心不足，现价下滑","#00FF00",RndID,-1
			elseif Rnum<90 then
				RndEvent 1000-Rnum*10,10,rs("QiYe"),"放出非利好消息，现价下滑","#00FF00",RndID,-1
			else
				RndEvent 1000-Rnum*10,1,rs("QiYe"),"放出非利好消息，现价微跌","#00FF00",RndID,-1
			end if			
		end if
	end if '2
	rs.close
end if '随机事件结束
if AI_Setting(0)=1 then call AI()		'如果打开机器人事件，则调用机器人
if Gupiao_Setting(12)>0 then
	'只保留最新Gupiao_Setting(12)个事件
	gp_conn.execute("delete from RndEvent where id not in(select top "&Gupiao_Setting(12)&" id from RndEvent order by id desc)")
end if
'===========================================================================
' Rate-价格变化的可能最大值，obj－对象，Flag－涨跌标记 -1 表示跌 1表示升
sub RndEvent(Rmax,Rmin,obj,message,msgColor,GPid,Flag)
	dim content,AMP,rs
	AMP=formatnumber(((Rmax-Rmin+1)*rnd+Rmin)/100,2,true)
	set rs=gp_conn.execute("select * from [GuPiao] where cid="&GPid&"")
	if rs("DangQianJiaGe")*(1+AMP*Flag/100)>rs("KaiPanJiaGe")*(1+Trade_Setting(9)) then
		AMP=formatnumber((rs("KaiPanJiaGe")*(1+Trade_Setting(9))-rs("DangQianJiaGe"))/rs("DangQianJiaGe"),2,true)
	elseif rs("DangQianJiaGe")*(1+AMP*Flag/100)<rs("KaiPanJiaGe")*(1-Trade_Setting(10)) then
		AMP=formatnumber((rs("DangQianJiaGe")-rs("KaiPanJiaGe")*(1-Trade_Setting(10)))/rs("DangQianJiaGe"),2,true)
	end if
	content="<font color=""#FFFF00"">"&obj&"</font><font color="""&msgColor&""">"&message&" "&AMP&" %</font>"
	gp_conn.execute("insert into RndEvent(content,addtime) values('"&content&"',now())") 
	gp_conn.execute("update [GuPiao] set DangQianJiaGe=DangQianJiaGe*"&(1+AMP*Flag/100)&",TodayWave=TodayWave+"&AMP*Flag&",TotalWave=TotalWave+"&AMP*Flag&" where Cid="&GPid&"")
end sub
'=============================================================================
sub AI()
	Randomize
	Rnum=rnd
	if Rnum>=1-AI_Setting(1) then
		dim RndID,StockNum,StockName,StockPrice
		dim BuyNum,OpenPrice
		RndID=gp_conn.execute("select top 1 cid from [GuPiao] order by cid desc")(0)		'现有股票的数目  
		Randomize Rnd
		RndID=int(RndID*rnd)+1		'被AI操作的股票ID
		set rs=gp_conn.execute("select JiaoYiLiang,QiYe,DangQianJiaGe,KaiPanJiaGe from [GuPiao] where Cid="&RndID)
		StockNum=rs(0) :	StockName=rs(1) : StockPrice=rs(2) : OpenPrice=rs(3)
		rs.close
		if StockNum+0<100 then exit sub		'当股票的交易量小于100时不进行操作
		'Randomize
		Rnum=int(rnd*10+1)
		select case Rnum
			case 1		'AI买入股票
				BuyNum=int(rnd*10+1)
				call AI_Buy(BuyNum,"机器人",-1,RndID,StockName,StockPrice,StockNum,OpenPrice)
			case 2
				BuyNum=int(rnd*100+1)
				call AI_Buy(BuyNum,"机器人",-1,RndID,StockName,StockPrice,StockNum,OpenPrice)
			case 4
				BuyNum=int(rnd*50+1)
				call AI_Buy(BuyNum,"机器人",-1,RndID,StockName,StockPrice,StockNum,OpenPrice)
			case 7
				BuyNum=int(rnd*65+1)
				call AI_Buy(BuyNum,"机器人",-1,RndID,StockName,StockPrice,StockNum,OpenPrice)
			case 9
				BuyNum=int(rnd*5+1)
				call AI_Buy(BuyNum,"机器人",-1,RndID,StockName,StockPrice,StockNum,OpenPrice)
			case else
				call AI_Sale("机器人",-1)
		end select
	elseif Rnum<0.10 then
		gp_conn.execute("insert into RndEvent(content,addtime) values('<font color=white>机器人</font><font color=#33CC99>在股票市场转了一圈，一笔交易都没有做成</font>',now())") 
	end if
end sub

sub AI_Buy(BuyNum,AIName,Uid,Sid,GupiaoName,NowPrice,StockNum,OpenPrice)
	Dim NewPrice,Zd		'股票新价格,涨跌幅
	set rs=gp_conn.execute("select ChiGuShu,MaiRuJiaGe from [DaHu] where sid="&sid&" and uid="&Uid&"")
	if rs.eof and rs.bof then	'1		'如果没有买过这个股票
		Randomize
		Rnum=Rnd
		if Rnum<=0.4 then
			Rnum=Rnum+0.5
		end if
		
		NewPrice=NowPrice*(1+BuyNum*Rnum/(StockNum*5)) 		'股票新价格
		if NewPrice>OpenPrice*(1+Trade_Setting(9)) then
			NewPrice=OpenPrice*(1+Trade_Setting(9))
		elseif NewPrice<OpenPrice*(1-Trade_Setting(10)) then
			NewPrice=OpenPrice*(1-Trade_Setting(10))
		end if
		sql="update [GuPiao] set DangQianJiaGe="&NewPrice&", JiaoYiLiang=JiaoYiLiang-"&BuyNum&",ShengYuGuFen=ShengYuGuFen-"&BuyNum&",ChengJiaoLiang=ChengJiaoLiang+"&BuyNum&",MaiRuBiShu=MaiRuBiShu+1 where sid="&sid 
		gp_conn.execute sql
		sql="insert into [DaHu] (Uid,ZhangHao,sid,MaiRuJiaGe,PingJunJiaGe,ChiGuShu,QiYe) values ("&Uid&",'"&AIName&"','"&sid&"','"&NowPrice&"','"&NowPrice&"','"&BuyNum&"','"&GupiaoName&"')"
		gp_conn.execute sql
		sql="update ai set cash=cash-"&BuyNum*NowPrice&"*1.01,fund=fund+"&BuyNum*NowPrice&"*0.99,StockTypeNum=StockTypeNum+1,TodayBuy=TodayBuy+1,ToTalBuy=ToTalBuy+1 where  AID="&Uid
		gp_conn.execute sql
		
		ZD=(NewPrice-NowPrice)/NowPrice
		mess="<font color=white>"&AIName&"</font><font color=#CC66FF>买入"&GupiaoName&" "&BuyNum&" 股，现价上涨 "&formatpercent(Zd,2,-1)&"</font>"
		sql="insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )"
		gp_conn.execute sql  
	else
		dim MyNum,AgvPrice	'持股数，平均价格
		MyNum=rs("ChiGuShu")
		
		Randomize
		Rnum=Rnd
		if Rnum<=0.4 then
			Rnum=Rnum+0.5
		end if
		NewPrice=NowPrice*(1+BuyNum*Rnum/(StockNum*5)) 		'股票新价格
		if NewPrice>OpenPrice*(1+Trade_Setting(9)) then
			NewPrice=OpenPrice*(1+Trade_Setting(9))
		elseif NewPrice<OpenPrice*(1-Trade_Setting(10)) then
			NewPrice=OpenPrice*(1-Trade_Setting(10))
		end if
		AgvPrice=(NowPrice*BuyNum+rs("MaiRuJiaGe")*MyNum)/(BuyNum+MyNum)  	'平均价格			
		
		sql="update [GuPiao] set DangQianJiaGe="&NewPrice&", JiaoYiLiang=JiaoYiLiang-"&BuyNum&",ShengYuGuFen=ShengYuGuFen-"&BuyNum&",ChengJiaoLiang=ChengJiaoLiang+"&BuyNum&",MaiRuBiShu=MaiRuBiShu+1 where sid="&sid
		gp_conn.execute sql
		sql="update [DaHu] set PingJunJiaGe="&AgvPrice&",MaiRuJiaGe="&NowPrice&",ChiGuShu=ChiGuShu+"&BuyNum&" where Uid="&Uid&" and sid="&sid
		gp_conn.execute sql
		sql="update ai set cash=cash-"&BuyNum*NowPrice&"*1.01,fund=fund+"&BuyNum*NowPrice&"*0.99,TodayBuy=TodayBuy+1,ToTalBuy=ToTalBuy+1 where  AID="&Uid
		gp_conn.execute sql
		
		ZD=(NewPrice-NowPrice)/NowPrice
		mess="<font color=white>"&AIName&"</font><font color=#FF00FF>再次买入"&GupiaoName&" "&BuyNum&" 股，现价上涨 "&formatpercent(Zd,2,-1)&"</font>"
		sql="insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"')"
		gp_conn.execute sql  
	end if
	rs.close
	gp_conn.execute("update [GupiaoConfig] set TodayBuy=TodayBuy+1,TodayTotal=TodayTotal+"&BuyNum*NowPrice&"")
end sub	

sub AI_Sale(AIName,AIID)
	dim rst,SaleNum,NowPrice,GupiaoName,Sid,NewPrice,OpenPrice	'  , 卖出数量  ， 股票现价 ， 股票名字 ， 股票ID
	set rst=server.createobject("adodb.recordset")
	sql="select d.ChiGuShu,g.DangQianJiaGe,g.QiYe,d.sid,g.JiaoYiLiang,g.KaiPanJiaGe from [DaHu] d inner join [GuPiao] g on d.sid=g.sid where d.Uid="&AIID
	rst.open sql,gp_conn,1,1
	if rst.eof and rst.bof then exit sub
	Randomize
	rst.absolutePosition=Fix(rst.recordcount*Rnd)+1
	SaleNum=rst(0)	'持股数
	NowPrice=rst(1)
	GupiaoName=rst(2)
	Sid=rst(3)
	OpenPrice=rst(5)
	if rnd<0.3 then		'机器人全部卖出某个股票
		Randomize
		Rnum=Rnd
		if Rnum<=0.4 then Rnum=Rnum+0.5
		NewPrice=NowPrice*(1-SaleNum*Rnum/(rst(4)*10))
		if NewPrice>OpenPrice*(1+Trade_Setting(9)) then
			NewPrice=OpenPrice*(1+Trade_Setting(9))
		elseif NewPrice<OpenPrice*(1-Trade_Setting(10)) then
			NewPrice=OpenPrice*(1-Trade_Setting(10))
		end if
		gp_conn.execute "update [GuPiao] set ChengJiaoLiang=ChengJiaoLiang+"&SaleNum&",DangQianJiaGe="&NewPrice&", JiaoYiLiang=JiaoYiLiang+"&SaleNum&",ShengYuGuFen=ShengYuGuFen+"&SaleNum&",MaiChuBiShu=MaiChuBiShu+1 where sid="&sid
		gp_conn.execute "delete from [DaHu] where Uid="&AIID&" and sid="&sid
		gp_conn.execute "update [ai] set cash=cash+"&NowPrice*SaleNum&"*0.99,fund=fund-"&NowPrice*SaleNum&"*1.01,StockTypeNum=StockTypeNum-1,TodaySale=TodaySale+1,ToTalSale=ToTalSale+1 where AID="&AIID
	
		mess="<font color=white>"&AIName&"</font><font color=#33CCFF>卖出"&GupiaoName&" "&SaleNum&" 股，现价下滑 "&formatpercent((NowPrice-NewPrice)/NowPrice,2,-1)&"</font>"
		gp_conn.execute "insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )"  
	else
		Randomize
		Rnum=Rnd
		if Rnum<=0.4 then Rnum=Rnum+0.5
		NewPrice=NowPrice*(1-SaleNum*Rnum/(rst(4)*10))		'卖出后股票的价格
		if NewPrice>OpenPrice*(1+Trade_Setting(9)) then
			NewPrice=OpenPrice*(1+Trade_Setting(9))
		elseif NewPrice<OpenPrice*(1-Trade_Setting(10)) then
			NewPrice=OpenPrice*(1-Trade_Setting(10))
		end if
		Rnum=Rnd
		if Rnum>0.5 then Rnum=1-Rnum   
		SaleNum=int(SaleNum*Rnum+1)		' 卖出股票的数量：1～自己的持股数量*0.5  的随机数	
		gp_conn.execute "update [GuPiao] set ChengJiaoLiang=ChengJiaoLiang+"&SaleNum&",DangQianJiaGe="&NewPrice&", JiaoYiLiang=JiaoYiLiang+"&SaleNum&",ShengYuGuFen=ShengYuGuFen+"&SaleNum&",MaiChuBiShu=MaiChuBiShu+1 where sid="&sid
		gp_conn.execute "update [DaHu] set ChiGuShu=ChiGuShu-"&SaleNum&" where Uid="&AIID&" and sid="&sid
		gp_conn.execute "update [ai] set cash=cash+"&NowPrice*SaleNum&"*0.99,fund=fund-"&NowPrice*SaleNum&"*1.01,TodaySale=TodaySale+1,ToTalSale=ToTalSale+1 where AID="&AIID
	
		mess="<font color=white>"&AIName&"</font><font color=#33CCFF>卖出"&GupiaoName&" "&SaleNum&" 股，现价下滑 "&formatpercent((NowPrice-NewPrice)/NowPrice,2,-1)&"</font>"
		gp_conn.execute "insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )"  
	end if
	rst.close
	set rst=nothing
	gp_conn.execute("update [GupiaoConfig] set TodaySale=TodaySale+1,TodayTotal=TodayTotal+"&NowPrice*SaleNum&"")
end sub	
%>