<%dim Rnum
Randomize
Rnum=rnd
if Rnum<=Gupiao_Setting(10)+0 then '�벻Ҫ�޸ı��䣬������ܻ����
	dim RndID,mess,MaxID
	MaxID=gp_conn.execute("select top 1 cid from [GuPiao] order by cid desc")(0)		'MaxID �������й�Ʊ����Ŀ  
	Randomize rnd
	RndID=int(MaxID*rnd)+1		'����¼�����������¼���Ӧ�Ĺ�Ʊ��CID
	sql= "select * from [GuPiao] where Cid="&RndID        
	set rs=gp_conn.execute(sql)
	if (rs("DangQianJiaGe")-rs("KaiPanJiaGe"))/rs("KaiPanJiaGe")>=Trade_Setting(9)+0 then	'��ͣ
		sql="update [GuPiao] set DangQianJiaGe=DangQianJiaGe*0.93 where Cid="&RndID&""
		gp_conn.execute sql           
		mess="<font color=#FFFF00>"&rs("QiYe")&"</font><font color=white>�Ĺ�Ʊ������ۣ�����ͣ״̬��</font>"
		sql="insert into RndEvent(content,addtime) values('"&mess&"',now())"
		gp_conn.execute sql  
	elseif (rs("DangQianJiaGe")-rs("KaiPanJiaGe"))/rs("KaiPanJiaGe")<=-Trade_Setting(10)+0 then	'��ͣ	
		sql="update [GuPiao] set DangQianJiaGe=DangQianJiaGe*1.08 where Cid="&RndID&""
		gp_conn.execute sql              
		mess="<font color=#FFFF00>"&rs("QiYe")&"</font><font color=white>�Ĺ�Ʊ���������룬�򿪵�ͣ״̬��</font>"
		sql="insert into RndEvent(content,addtime) values('"&mess&"',now())"
		gp_conn.execute sql  
	else   
		Randomize
		if Rnd<Gupiao_Setting(11)+0 then		'��Ʊ�۸����� �¼�
			'Randomize
			Rnum=formatnumber(100*rnd+1,2,true)		'1��100 ֮��������
			if Rnum<10 then 
				RndEvent Rnum*10,1,rs("QiYe"),"�ʲ����飬�ּ�΢��","red",RndID,1 
			elseif Rnum<20 then
				RndEvent Rnum*10,1,rs("QiYe"),"�ʲ����飬�ּ�����","red",RndID,1
			elseif Rnum<30 then
				RndEvent Rnum*10,10,rs("QiYe"),"��˾ҵ����ߣ�Ͷ�������ļӴ��ּ�����","red",RndID,1
			elseif Rnum<40 then
				RndEvent Rnum*10,20,rs("QiYe"),"���ģ�ʲ����飬�ּ�����","red",RndID,1
			elseif Rnum<50 then
				RndEvent Rnum*10,300,rs("QiYe"),"ҵ�������ߣ�ע������ʽ��ּ�����","red",RndID,1
			elseif Rnum<60 then 
				RndEvent 1000-Rnum*10,40,rs("QiYe"),"ҵ�������ߣ�ע������ʽ��ּ�����","red",RndID,1 
			elseif Rnum<70 then
				RndEvent 1000-Rnum*10,30,rs("QiYe"),"���ģ�ʲ����飬�ּ�����","red",RndID,1
			elseif Rnum<80 then
				RndEvent 1000-Rnum*10,20,rs("QiYe"),"��˾ҵ����ߣ�Ͷ�������ļӴ��ּ�����","red",RndID,1
			elseif Rnum<90 then
				RndEvent 1000-Rnum*10,10,rs("QiYe"),"�ʲ����飬�ּ�����","red",RndID,1
			else
				RndEvent 1000-Rnum*10,1,rs("QiYe"),"�ʲ����飬�ּ�΢��","red",RndID,1
			end if
		else									'��Ʊ�۸��»� �¼�
			'Randomize
			Rnum=formatnumber(100*rnd+1,2,true)		'1��100 ֮��������  
			if Rnum<10 then 
				RndEvent Rnum*10,1,rs("QiYe"),"�ų���������Ϣ���ּ�΢��","#00FF00",RndID,-1 
			elseif Rnum<20 then
				RndEvent Rnum*10,1,rs("QiYe"),"�ų���������Ϣ���ּ��»�","#00FF00",RndID,-1
			elseif Rnum<30 then
				RndEvent Rnum*10,10,rs("QiYe"),"��˾ҵ�����ѣ�Ͷ�������Ĳ��㣬�ּ��»�","#00FF00",RndID,-1
			elseif Rnum<40 then
				RndEvent Rnum*10,20,rs("QiYe"),"ҵ����ʼ���£��ּ��µ�","#00FF00",RndID,-1
			elseif Rnum<50 then
				RndEvent Rnum*10,300,rs("QiYe"),"ҵ��������£��ʽ��㣬�ּ��µ�","#00FF00",RndID,-1
			elseif Rnum<60 then 
				RndEvent 1000-Rnum*10,40,rs("QiYe"),"ҵ��������£����ܴ����ۣ��ּ��µ�","#00FF00",RndID,-1 
			elseif Rnum<70 then
				RndEvent 1000-Rnum*10,30,rs("QiYe"),"ҵ����ʼ���£��ּ��µ�","#00FF00",RndID,-1
			elseif Rnum<80 then
				RndEvent 1000-Rnum*10,20,rs("QiYe"),"��˾ҵ�����ѣ�Ͷ�������Ĳ��㣬�ּ��»�","#00FF00",RndID,-1
			elseif Rnum<90 then
				RndEvent 1000-Rnum*10,10,rs("QiYe"),"�ų���������Ϣ���ּ��»�","#00FF00",RndID,-1
			else
				RndEvent 1000-Rnum*10,1,rs("QiYe"),"�ų���������Ϣ���ּ�΢��","#00FF00",RndID,-1
			end if			
		end if
	end if '2
	rs.close
end if '����¼�����
if AI_Setting(0)=1 then call AI()		'����򿪻������¼�������û�����
if Gupiao_Setting(12)>0 then
	'ֻ��������Gupiao_Setting(12)���¼�
	gp_conn.execute("delete from RndEvent where id not in(select top "&Gupiao_Setting(12)&" id from RndEvent order by id desc)")
end if
'===========================================================================
' Rate-�۸�仯�Ŀ������ֵ��obj������Flag���ǵ���� -1 ��ʾ�� 1��ʾ��
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
		RndID=gp_conn.execute("select top 1 cid from [GuPiao] order by cid desc")(0)		'���й�Ʊ����Ŀ  
		Randomize Rnd
		RndID=int(RndID*rnd)+1		'��AI�����Ĺ�ƱID
		set rs=gp_conn.execute("select JiaoYiLiang,QiYe,DangQianJiaGe,KaiPanJiaGe from [GuPiao] where Cid="&RndID)
		StockNum=rs(0) :	StockName=rs(1) : StockPrice=rs(2) : OpenPrice=rs(3)
		rs.close
		if StockNum+0<100 then exit sub		'����Ʊ�Ľ�����С��100ʱ�����в���
		'Randomize
		Rnum=int(rnd*10+1)
		select case Rnum
			case 1		'AI�����Ʊ
				BuyNum=int(rnd*10+1)
				call AI_Buy(BuyNum,"������",-1,RndID,StockName,StockPrice,StockNum,OpenPrice)
			case 2
				BuyNum=int(rnd*100+1)
				call AI_Buy(BuyNum,"������",-1,RndID,StockName,StockPrice,StockNum,OpenPrice)
			case 4
				BuyNum=int(rnd*50+1)
				call AI_Buy(BuyNum,"������",-1,RndID,StockName,StockPrice,StockNum,OpenPrice)
			case 7
				BuyNum=int(rnd*65+1)
				call AI_Buy(BuyNum,"������",-1,RndID,StockName,StockPrice,StockNum,OpenPrice)
			case 9
				BuyNum=int(rnd*5+1)
				call AI_Buy(BuyNum,"������",-1,RndID,StockName,StockPrice,StockNum,OpenPrice)
			case else
				call AI_Sale("������",-1)
		end select
	elseif Rnum<0.10 then
		gp_conn.execute("insert into RndEvent(content,addtime) values('<font color=white>������</font><font color=#33CC99>�ڹ�Ʊ�г�ת��һȦ��һ�ʽ��׶�û������</font>',now())") 
	end if
end sub

sub AI_Buy(BuyNum,AIName,Uid,Sid,GupiaoName,NowPrice,StockNum,OpenPrice)
	Dim NewPrice,Zd		'��Ʊ�¼۸�,�ǵ���
	set rs=gp_conn.execute("select ChiGuShu,MaiRuJiaGe from [DaHu] where sid="&sid&" and uid="&Uid&"")
	if rs.eof and rs.bof then	'1		'���û����������Ʊ
		Randomize
		Rnum=Rnd
		if Rnum<=0.4 then
			Rnum=Rnum+0.5
		end if
		
		NewPrice=NowPrice*(1+BuyNum*Rnum/(StockNum*5)) 		'��Ʊ�¼۸�
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
		mess="<font color=white>"&AIName&"</font><font color=#CC66FF>����"&GupiaoName&" "&BuyNum&" �ɣ��ּ����� "&formatpercent(Zd,2,-1)&"</font>"
		sql="insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )"
		gp_conn.execute sql  
	else
		dim MyNum,AgvPrice	'�ֹ�����ƽ���۸�
		MyNum=rs("ChiGuShu")
		
		Randomize
		Rnum=Rnd
		if Rnum<=0.4 then
			Rnum=Rnum+0.5
		end if
		NewPrice=NowPrice*(1+BuyNum*Rnum/(StockNum*5)) 		'��Ʊ�¼۸�
		if NewPrice>OpenPrice*(1+Trade_Setting(9)) then
			NewPrice=OpenPrice*(1+Trade_Setting(9))
		elseif NewPrice<OpenPrice*(1-Trade_Setting(10)) then
			NewPrice=OpenPrice*(1-Trade_Setting(10))
		end if
		AgvPrice=(NowPrice*BuyNum+rs("MaiRuJiaGe")*MyNum)/(BuyNum+MyNum)  	'ƽ���۸�			
		
		sql="update [GuPiao] set DangQianJiaGe="&NewPrice&", JiaoYiLiang=JiaoYiLiang-"&BuyNum&",ShengYuGuFen=ShengYuGuFen-"&BuyNum&",ChengJiaoLiang=ChengJiaoLiang+"&BuyNum&",MaiRuBiShu=MaiRuBiShu+1 where sid="&sid
		gp_conn.execute sql
		sql="update [DaHu] set PingJunJiaGe="&AgvPrice&",MaiRuJiaGe="&NowPrice&",ChiGuShu=ChiGuShu+"&BuyNum&" where Uid="&Uid&" and sid="&sid
		gp_conn.execute sql
		sql="update ai set cash=cash-"&BuyNum*NowPrice&"*1.01,fund=fund+"&BuyNum*NowPrice&"*0.99,TodayBuy=TodayBuy+1,ToTalBuy=ToTalBuy+1 where  AID="&Uid
		gp_conn.execute sql
		
		ZD=(NewPrice-NowPrice)/NowPrice
		mess="<font color=white>"&AIName&"</font><font color=#FF00FF>�ٴ�����"&GupiaoName&" "&BuyNum&" �ɣ��ּ����� "&formatpercent(Zd,2,-1)&"</font>"
		sql="insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"')"
		gp_conn.execute sql  
	end if
	rs.close
	gp_conn.execute("update [GupiaoConfig] set TodayBuy=TodayBuy+1,TodayTotal=TodayTotal+"&BuyNum*NowPrice&"")
end sub	

sub AI_Sale(AIName,AIID)
	dim rst,SaleNum,NowPrice,GupiaoName,Sid,NewPrice,OpenPrice	'  , ��������  �� ��Ʊ�ּ� �� ��Ʊ���� �� ��ƱID
	set rst=server.createobject("adodb.recordset")
	sql="select d.ChiGuShu,g.DangQianJiaGe,g.QiYe,d.sid,g.JiaoYiLiang,g.KaiPanJiaGe from [DaHu] d inner join [GuPiao] g on d.sid=g.sid where d.Uid="&AIID
	rst.open sql,gp_conn,1,1
	if rst.eof and rst.bof then exit sub
	Randomize
	rst.absolutePosition=Fix(rst.recordcount*Rnd)+1
	SaleNum=rst(0)	'�ֹ���
	NowPrice=rst(1)
	GupiaoName=rst(2)
	Sid=rst(3)
	OpenPrice=rst(5)
	if rnd<0.3 then		'������ȫ������ĳ����Ʊ
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
	
		mess="<font color=white>"&AIName&"</font><font color=#33CCFF>����"&GupiaoName&" "&SaleNum&" �ɣ��ּ��»� "&formatpercent((NowPrice-NewPrice)/NowPrice,2,-1)&"</font>"
		gp_conn.execute "insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )"  
	else
		Randomize
		Rnum=Rnd
		if Rnum<=0.4 then Rnum=Rnum+0.5
		NewPrice=NowPrice*(1-SaleNum*Rnum/(rst(4)*10))		'�������Ʊ�ļ۸�
		if NewPrice>OpenPrice*(1+Trade_Setting(9)) then
			NewPrice=OpenPrice*(1+Trade_Setting(9))
		elseif NewPrice<OpenPrice*(1-Trade_Setting(10)) then
			NewPrice=OpenPrice*(1-Trade_Setting(10))
		end if
		Rnum=Rnd
		if Rnum>0.5 then Rnum=1-Rnum   
		SaleNum=int(SaleNum*Rnum+1)		' ������Ʊ��������1���Լ��ĳֹ�����*0.5  �������	
		gp_conn.execute "update [GuPiao] set ChengJiaoLiang=ChengJiaoLiang+"&SaleNum&",DangQianJiaGe="&NewPrice&", JiaoYiLiang=JiaoYiLiang+"&SaleNum&",ShengYuGuFen=ShengYuGuFen+"&SaleNum&",MaiChuBiShu=MaiChuBiShu+1 where sid="&sid
		gp_conn.execute "update [DaHu] set ChiGuShu=ChiGuShu-"&SaleNum&" where Uid="&AIID&" and sid="&sid
		gp_conn.execute "update [ai] set cash=cash+"&NowPrice*SaleNum&"*0.99,fund=fund-"&NowPrice*SaleNum&"*1.01,TodaySale=TodaySale+1,ToTalSale=ToTalSale+1 where AID="&AIID
	
		mess="<font color=white>"&AIName&"</font><font color=#33CCFF>����"&GupiaoName&" "&SaleNum&" �ɣ��ּ��»� "&formatpercent((NowPrice-NewPrice)/NowPrice,2,-1)&"</font>"
		gp_conn.execute "insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )"  
	end if
	rst.close
	set rst=nothing
	gp_conn.execute("update [GupiaoConfig] set TodaySale=TodaySale+1,TodayTotal=TodayTotal+"&NowPrice*SaleNum&"")
end sub	
%>