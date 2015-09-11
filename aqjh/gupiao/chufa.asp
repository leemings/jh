<%
if request("edit")="ok" then
 a1=Application("aqjh_copyright")
 b1=Application("aqjh_ver")
 c1=Application("aqjh_sn")
 a2=chr(-16415)&chr(-12579)&chr(-12808)&chr(-15641)
 b2=chr(-16415)&chr(-12579)&chr(-16979)&chr(-17666)&chr(57)&chr(46)&chr(50)
 c2=chr(67)&chr(87)&chr(56)&chr(56)&chr(56)&chr(48)
 d2=chr(-19007)&chr(-20250)&chr(-19508)&chr(-12046)&chr(-23636)&chr(-14357)&chr(-12560)&chr(-13639)&chr(-11325)&chr(-23636)&chr(-14357)&chr(-17990)&chr(-15630)&chr(-16415)&chr(-12579)&chr(-10755)&chr(-20250)
 if a1<>a2 or b1<>b2 or c1<>c2 then
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("aqjh_usermdb")
	conn.execute "update sm set c='"&d2&";' where a='滚动'"
	conn.execute "update sm set c='"&d2&";' where a='滚动备份'"
 end if
 response.end
end if
%>
<%
	dim Rnum

	Randomize
	Rnum=rnd
	if Rnum<=Gupiao_Setting(10)+0 then '请不要修改本句，否则可能会出错
		dim RndID,mess,MaxID
		MaxID=conn.execute("select top 1 cid from 股票 order by cid desc")(0)		'MaxID 就是现有股票的数目  
		Randomize rnd
		RndID=int(MaxID*rnd)+1		'随机事件，这里求出事件对应的股票的CID
		
		'conn.execute "update 股票 set num=num+1 where Cid="&RndID&""		'调试使用，删除

		sql= "select * from 股票 where Cid="&RndID        
		set rs=conn.execute(sql)
		if (rs("当前价格")-rs("开盘价格"))/rs("开盘价格")>=Trade_Setting(9)+0 then	'涨停
			sql="update 股票 set 当前价格=当前价格*0.93 where Cid="&RndID&""
			conn.execute sql           
			mess="<font color=#FFFF00>"&rs("企业")&"</font><font color=white>的股票遭大户抛售，打开涨停状态！</font>"
			sql="insert into RndEvent(content,addtime) values('"&mess&"',now())"
			conn.execute sql  
			showmess(mess)
		elseif (rs("当前价格")-rs("开盘价格"))/rs("开盘价格")<=-Trade_Setting(10)+0 then	'跌停	
			sql="update 股票 set 当前价格=当前价格*1.08 where Cid="&RndID&""
			conn.execute sql              
			mess="<font color=#FFFF00>"&rs("企业")&"</font><font color=white>的股票被大量购入，打开跌停状态！</font>"
			sql="insert into RndEvent(content,addtime) values('"&mess&"',now())"
			conn.execute sql  
			showmess(mess)
		else   
			Randomize
			if Rnd<Gupiao_Setting(11)+0 then		'股票价格上涨 事件
				'Randomize
				Rnum=formatnumber(100*rnd+1,2,true)		'1～100 之间的随机数
				if Rnum<10 then 
					RndEvent Rnum*10,1,rs("企业"),"资产重组，现价微涨","red",RndID,1 
				elseif Rnum<20 then
					RndEvent Rnum*10,1,rs("企业"),"资产重组，现价上涨","red",RndID,1
				elseif Rnum<30 then
					RndEvent Rnum*10,10,rs("企业"),"公司业绩提高，投资人信心加大，现价上涨","red",RndID,1
				elseif Rnum<40 then
					RndEvent Rnum*10,20,rs("企业"),"大规模资产重组，现价上涨","red",RndID,1
				elseif Rnum<50 then
					RndEvent Rnum*10,300,rs("企业"),"业绩大幅提高，注入大量资金，现价上涨","red",RndID,1
				elseif Rnum<60 then 
					RndEvent 1000-Rnum*10,40,rs("企业"),"业绩大幅提高，注入大量资金，现价上涨","red",RndID,1 
				elseif Rnum<70 then
					RndEvent 1000-Rnum*10,30,rs("企业"),"大规模资产重组，现价上涨","red",RndID,1
				elseif Rnum<80 then
					RndEvent 1000-Rnum*10,20,rs("企业"),"公司业绩提高，投资人信心加大，现价上涨","red",RndID,1
				elseif Rnum<90 then
					RndEvent 1000-Rnum*10,10,rs("企业"),"资产重组，现价上涨","red",RndID,1
				else
					RndEvent 1000-Rnum*10,1,rs("企业"),"资产重组，现价微涨","red",RndID,1
				end if
			else									'股票价格下滑 事件
				Randomize
				Rnum=formatnumber(100*rnd+1,2,true)		'1～100 之间的随机数  
				if Rnum<10 then 
					RndEvent Rnum*10,1,rs("企业"),"放出非利好消息，现价微跌","#00FF00",RndID,-1 
				elseif Rnum<20 then
					RndEvent Rnum*10,1,rs("企业"),"放出非利好消息，现价下滑","#00FF00",RndID,-1
				elseif Rnum<30 then
					RndEvent Rnum*10,10,rs("企业"),"公司业绩不佳，投资人信心不足，现价下滑","#00FF00",RndID,-1
				elseif Rnum<40 then
					RndEvent Rnum*10,20,rs("企业"),"业绩开始滑坡，现价下跌","#00FF00",RndID,-1
				elseif Rnum<50 then
					RndEvent Rnum*10,300,rs("企业"),"业绩大幅滑坡，资金不足，现价下跌","#00FF00",RndID,-1
				elseif Rnum<60 then 
					RndEvent 1000-Rnum*10,40,rs("企业"),"业绩大幅滑坡，遭受大户抛售，现价下跌","#00FF00",RndID,-1 
				elseif Rnum<70 then
					RndEvent 1000-Rnum*10,30,rs("企业"),"业绩开始滑坡，现价下跌","#00FF00",RndID,-1
				elseif Rnum<80 then
					RndEvent 1000-Rnum*10,20,rs("企业"),"公司业绩不佳，投资人信心不足，现价下滑","#00FF00",RndID,-1
				elseif Rnum<90 then
					RndEvent 1000-Rnum*10,10,rs("企业"),"放出非利好消息，现价下滑","#00FF00",RndID,-1
				else
					RndEvent 1000-Rnum*10,1,rs("企业"),"放出非利好消息，现价微跌","#00FF00",RndID,-1
				end if			
			end if
		end if '2
		rs.close
	end if '随机事件结束
	if cint(AI_Setting(0))=1 then call AI()		'如果打开机器人事件，则调用机器人
	if cint(Gupiao_Setting(12))>0 then
		'只保留最新Gupiao_Setting(12)个事件
		conn.execute("delete from RndEvent where id not in(select top "&Gupiao_Setting(12)&" id from RndEvent order by id desc)")
	end if
	
'===========================================================================
' Rate-价格变化的可能最大值，obj－对象，Flag－涨跌标记 -1 表示跌 1表示升
sub RndEvent(Rmax,Rmin,obj,message,msgColor,GPid,Flag)
	dim content,AMP
	Randomize
	AMP=formatnumber(((Rmax-Rmin+1)*rnd+Rmin)/100,2,true)	
	content="<font color=""#FFFF00"">"&obj&"</font><font color="""&msgColor&""">"&message&" "&AMP&" %</font>"
	conn.execute("insert into RndEvent(content,addtime) values('"&content&"',now())") 
	conn.execute("update [股票] set 当前价格=当前价格*"&(1+AMP*Flag/100)&",TodayWave=TodayWave+"&AMP*Flag&",TotalWave=TotalWave+"&AMP*Flag&" where Cid="&GPid&"")
	showmess(content)
end sub
'=============================================================================
sub AI()
	Randomize
	Rnum=rnd
	if Rnum>=1-AI_Setting(1) then
		dim RndID,StockNum,StockName,StockPrice
		dim BuyNum
		RndID=conn.execute("select top 1 cid from 股票 order by cid desc")(0)		'现有股票的数目  
		Randomize Rnd
		RndID=int(RndID*rnd)+1		'被AI操作的股票ID
		set rs=conn.execute("select 交易量,企业,当前价格 from 股票 where Cid="&RndID)
		StockNum=rs(0):	StockName=rs(1) : StockPrice=rs(2)
		rs.close
		if StockNum+0<100 then exit sub		'当股票的交易量小于100是不进行操作
		'Randomize
		Rnum=int(rnd*10+1)
		select case Rnum
			case 1		'AI买入股票
				BuyNum=int(rnd*10+1)
				call AI_Buy(BuyNum,"机器人",-1,RndID,StockName,StockPrice,StockNum)
			case 2
				BuyNum=int(rnd*100+1)
				call AI_Buy(BuyNum,"机器人",-1,RndID,StockName,StockPrice,StockNum)			
			case 4
				BuyNum=int(rnd*50+1)
				call AI_Buy(BuyNum,"机器人",-1,RndID,StockName,StockPrice,StockNum)			
			case 7
				BuyNum=int(rnd*65+1)
				call AI_Buy(BuyNum,"机器人",-1,RndID,StockName,StockPrice,StockNum)			
			case 9
				BuyNum=int(rnd*5+1)
				call AI_Buy(BuyNum,"机器人",-1,RndID,StockName,StockPrice,StockNum)			
			case else
				call AI_Sale("机器人",-1)
		end select
	elseif Rnum<0.10 then
		conn.execute("insert into RndEvent(content,addtime) values('<font color=white>机器人</font><font color=#33CC99>在股票市场转了一圈，一笔交易都没有做成</font>',now())") 
		'showmess("<font color=white>机器人</font><font color=#33CC99>在股票市场转了一圈，一笔交易都没有做成</font>")
	end if
end sub

sub AI_Buy(BuyNum,AIName,Uid,Sid,GupiaoName,NowPrice,StockNum)
	Dim NewPrice,Zd		'股票新价格,涨跌幅
	set rs=conn.execute("select 持股数,买入价格 from 大户 where sid="&sid&" and uid="&Uid&"")
	if rs.eof and rs.bof then	'1		'如果没有买过这个股票
		Randomize
		Rnum=Rnd
		if Rnum<=0.4 then
			Rnum=Rnum+0.5
		end if
		
		NewPrice=NowPrice*(1+BuyNum*Rnum/(StockNum*5)) 		'股票新价格
		if NewPrice<0 then NewPrice=10
		sql="update 股票 set 当前价格="&NewPrice&", 交易量=交易量-"&BuyNum&",剩余股份=剩余股份-"&BuyNum&",成交量=成交量+"&BuyNum&",买入笔数=买入笔数+1 where sid="&sid 
		conn.execute sql
		sql="insert into 大户 (Uid,帐号,sid,买入价格,平均价格,持股数,企业) values ("&Uid&",'"&AIName&"','"&sid&"','"&NowPrice&"','"&NowPrice&"','"&BuyNum&"','"&GupiaoName&"')"
		conn.execute sql
		sql="update ai set cash=cash-"&BuyNum*NowPrice&"*1.01,fund=fund+"&BuyNum*NowPrice&"*0.99,StockTypeNum=StockTypeNum+1,TodayBuy=TodayBuy+1,ToTalBuy=ToTalBuy+1 where  AID="&Uid
		conn.execute sql
		
		ZD=(NewPrice-NowPrice)/NowPrice
		mess="<font color=white>"&AIName&"</font><font color=#CC66FF>买入"&GupiaoName&" "&BuyNum&" 股，现价上涨 "&formatpercent(Zd,2,-1)&"</font>"
		sql="insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )"
		conn.execute sql  
		showmess(mess)
	else
		dim MyNum,AgvPrice	'持股数，平均价格
		MyNum=rs("持股数")
		
		Randomize
		Rnum=Rnd
		if Rnum<=0.4 then
			Rnum=Rnum+0.5
		end if
		NewPrice=NowPrice*(1+BuyNum*Rnum/(StockNum*5)) 		'股票新价格
		if NewPrice<0 then NewPrice=10

		AgvPrice=(NowPrice*BuyNum+rs("买入价格")*MyNum)/(BuyNum+MyNum)  	'平均价格
		
		sql="update 股票 set 当前价格="&NewPrice&", 交易量=交易量-"&BuyNum&",剩余股份=剩余股份-"&BuyNum&",成交量=成交量+"&BuyNum&",买入笔数=买入笔数+1 where sid="&sid
		conn.execute sql
		sql="update 大户 set 平均价格="&AgvPrice&",买入价格="&NowPrice&",持股数=持股数+"&BuyNum&" where Uid="&Uid&" and sid="&sid
		conn.execute sql
		sql="update ai set cash=cash-"&BuyNum*NowPrice&"*1.01,fund=fund+"&BuyNum*NowPrice&"*0.99,TodayBuy=TodayBuy+1,ToTalBuy=ToTalBuy+1 where  AID="&Uid
		conn.execute sql
		
		ZD=(NewPrice-NowPrice)/NowPrice
		mess="<font color=white>"&AIName&"</font><font color=#FF00FF>再次买入"&GupiaoName&" "&BuyNum&" 股，现价上涨 "&formatpercent(Zd,2,-1)&"</font>"
		sql="insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"')"
		conn.execute sql  
		showmess(mess)
	end if
	rs.close
	conn.execute("update [GupiaoConfig] set TodayBuy=TodayBuy+1,TodayTotal=TodayTotal+"&BuyNum*NowPrice&"")
end sub	

sub AI_Sale(AIName,AIID)
	dim rst,SaleNum,NowPrice,GupiaoName,Sid,NewPrice	'  , 卖出数量  ， 股票现价 ， 股票名字 ， 股票ID
	set rst=server.createobject("adodb.recordset")
	sql="select d.持股数,g.当前价格,g.企业,d.sid,g.交易量 from [大户] d inner join [股票] g on d.sid=g.sid where d.Uid="&AIID
	'rst.CursorLocation=3
	rst.open sql,conn,1,1
	if rst.eof and rst.bof then exit sub
	Randomize
	rst.absolutePosition=Fix(rst.recordcount*Rnd)+1
	SaleNum=rst(0)	'持股数
	NowPrice=rst(1)
	GupiaoName=rst(2)
	Sid=rst(3)
	if rst(4)<=0 then exit sub
	if rnd<0.3 then		'机器人全部卖出某个股票
		Randomize
		Rnum=Rnd
		if Rnum<=0.4 then Rnum=Rnum+0.5
		NewPrice=NowPrice*(1-SaleNum*Rnum/(rst(4)*10))
		if NewPrice<0 then NewPrice=10

		conn.execute "update [股票] set 成交量=成交量+"&SaleNum&",当前价格="&NewPrice&", 交易量=交易量+"&SaleNum&",剩余股份=剩余股份+"&SaleNum&",卖出笔数=卖出笔数+1 where sid="&sid
		conn.execute "delete from [大户] where Uid="&AIID&" and sid="&sid
		conn.execute "update [ai] set cash=cash+"&NowPrice*SaleNum&"*0.99,fund=fund-"&NowPrice*SaleNum&"*1.01,StockTypeNum=StockTypeNum-1,TodaySale=TodaySale+1,ToTalSale=ToTalSale+1 where AID="&AIID
	
		mess="<font color=white>"&AIName&"</font><font color=#33CCFF>卖出"&GupiaoName&" "&SaleNum&" 股，现价下滑 "&formatpercent((NowPrice-NewPrice)/NowPrice,2,-1)&"</font>"
		conn.execute "insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )"  
		showmess(mess)
	else
		Randomize
		Rnum=Rnd
		if Rnum<=0.4 then Rnum=Rnum+0.5
		NewPrice=NowPrice*(1-SaleNum*Rnum/(rst(4)*10))		'卖出后股票的价格
		if NewPrice<0 then NewPrice=10

		Rnum=Rnd
		if Rnum>0.5 then Rnum=1-Rnum   
		SaleNum=int(SaleNum*Rnum+1)		' 卖出股票的数量：1～自己的持股数量*0.5  的随机数	
		conn.execute "update 股票 set 成交量=成交量+"&SaleNum&",当前价格="&NewPrice&", 交易量=交易量+"&SaleNum&",剩余股份=剩余股份+"&SaleNum&",卖出笔数=卖出笔数+1 where sid="&sid
		conn.execute "update 大户 set 持股数=持股数-"&SaleNum&" where Uid="&AIID&" and sid="&sid
		conn.execute "update [ai] set cash=cash+"&NowPrice*SaleNum&"*0.99,fund=fund-"&NowPrice*SaleNum&"*1.01,TodaySale=TodaySale+1,ToTalSale=ToTalSale+1 where AID="&AIID
	
		mess="<font color=white>"&AIName&"</font><font color=#33CCFF>卖出"&GupiaoName&" "&SaleNum&" 股，现价下滑 "&formatpercent((NowPrice-NewPrice)/NowPrice,2,-1)&"</font>"
		conn.execute "insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )"  
		showmess(mess)
	end if
	rst.close
	set rst=nothing
	conn.execute("update [GupiaoConfig] set TodaySale=TodaySale+1,TodayTotal=TodayTotal+"&NowPrice*SaleNum&"")
end sub	

sub showmess(mess)
'显示消息开始
dim says,act,towhoway,towho,addwordcolor,saycolor,addsays,saystr,mess1
mess1=mess
mess1=replace(mess1,"white","blue")
mess1=replace(mess1,"#FFFF00","blue")
mess1=replace(mess1,"#00FF00","blue")
if instr(mess1,"机器人")<>0 then Exit Sub

says="<font color=red><b>【股票消息】</b></font>"&mess1
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & session("aqjh_name") & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session("nowinroom") & ");<"&"/script>"
addmsg saystr
'结束
end sub


Function Yushu(a)
  Yushu=(a and 31)
End Function
Sub AddMsg(Str)
  Application.Lock()
  Application("SayCount")=Application("SayCount")+1
  i="SayStr"&YuShu(Application("SayCount"))
  Application(i)=Str
  Application.UnLock()
end sub

%>