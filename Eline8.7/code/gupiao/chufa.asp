<%
	dim Rnum

	Randomize
	Rnum=rnd
	if Rnum<=Gupiao_Setting(10)+0 then '�벻Ҫ�޸ı��䣬������ܻ����
		dim RndID,mess,MaxID
		MaxID=conn.execute("select top 1 cid from ��Ʊ order by cid desc")(0)		'MaxID �������й�Ʊ����Ŀ  
		Randomize rnd
		RndID=int(MaxID*rnd)+1		'����¼�����������¼���Ӧ�Ĺ�Ʊ��CID
		
		'conn.execute "update ��Ʊ set num=num+1 where Cid="&RndID&""		'����ʹ�ã�ɾ��

		sql= "select * from ��Ʊ where Cid="&RndID        
		set rs=conn.execute(sql)
		if (rs("��ǰ�۸�")-rs("���̼۸�"))/rs("���̼۸�")>=Trade_Setting(9)+0 then	'��ͣ
			sql="update ��Ʊ set ��ǰ�۸�=��ǰ�۸�*0.93 where Cid="&RndID&""
			conn.execute sql           
			mess="<font color=#FFFF00>"&rs("��ҵ")&"</font><font color=white>�Ĺ�Ʊ������ۣ�����ͣ״̬��</font>"
			sql="insert into RndEvent(content,addtime) values('"&mess&"',now())"
			conn.execute sql  
			showmess(mess)
		elseif (rs("��ǰ�۸�")-rs("���̼۸�"))/rs("���̼۸�")<=-Trade_Setting(10)+0 then	'��ͣ	
			sql="update ��Ʊ set ��ǰ�۸�=��ǰ�۸�*1.08 where Cid="&RndID&""
			conn.execute sql              
			mess="<font color=#FFFF00>"&rs("��ҵ")&"</font><font color=white>�Ĺ�Ʊ���������룬�򿪵�ͣ״̬��</font>"
			sql="insert into RndEvent(content,addtime) values('"&mess&"',now())"
			conn.execute sql  
			showmess(mess)
		else   
			Randomize
			if Rnd<Gupiao_Setting(11)+0 then		'��Ʊ�۸����� �¼�
				'Randomize
				Rnum=formatnumber(100*rnd+1,2,true)		'1��100 ֮��������
				if Rnum<10 then 
					RndEvent Rnum*10,1,rs("��ҵ"),"�ʲ����飬�ּ�΢��","red",RndID,1 
				elseif Rnum<20 then
					RndEvent Rnum*10,1,rs("��ҵ"),"�ʲ����飬�ּ�����","red",RndID,1
				elseif Rnum<30 then
					RndEvent Rnum*10,10,rs("��ҵ"),"��˾ҵ����ߣ�Ͷ�������ļӴ��ּ�����","red",RndID,1
				elseif Rnum<40 then
					RndEvent Rnum*10,20,rs("��ҵ"),"���ģ�ʲ����飬�ּ�����","red",RndID,1
				elseif Rnum<50 then
					RndEvent Rnum*10,300,rs("��ҵ"),"ҵ�������ߣ�ע������ʽ��ּ�����","red",RndID,1
				elseif Rnum<60 then 
					RndEvent 1000-Rnum*10,40,rs("��ҵ"),"ҵ�������ߣ�ע������ʽ��ּ�����","red",RndID,1 
				elseif Rnum<70 then
					RndEvent 1000-Rnum*10,30,rs("��ҵ"),"���ģ�ʲ����飬�ּ�����","red",RndID,1
				elseif Rnum<80 then
					RndEvent 1000-Rnum*10,20,rs("��ҵ"),"��˾ҵ����ߣ�Ͷ�������ļӴ��ּ�����","red",RndID,1
				elseif Rnum<90 then
					RndEvent 1000-Rnum*10,10,rs("��ҵ"),"�ʲ����飬�ּ�����","red",RndID,1
				else
					RndEvent 1000-Rnum*10,1,rs("��ҵ"),"�ʲ����飬�ּ�΢��","red",RndID,1
				end if
			else									'��Ʊ�۸��»� �¼�
				Randomize
				Rnum=formatnumber(100*rnd+1,2,true)		'1��100 ֮��������  
				if Rnum<10 then 
					RndEvent Rnum*10,1,rs("��ҵ"),"�ų���������Ϣ���ּ�΢��","#00FF00",RndID,-1 
				elseif Rnum<20 then
					RndEvent Rnum*10,1,rs("��ҵ"),"�ų���������Ϣ���ּ��»�","#00FF00",RndID,-1
				elseif Rnum<30 then
					RndEvent Rnum*10,10,rs("��ҵ"),"��˾ҵ�����ѣ�Ͷ�������Ĳ��㣬�ּ��»�","#00FF00",RndID,-1
				elseif Rnum<40 then
					RndEvent Rnum*10,20,rs("��ҵ"),"ҵ����ʼ���£��ּ��µ�","#00FF00",RndID,-1
				elseif Rnum<50 then
					RndEvent Rnum*10,300,rs("��ҵ"),"ҵ��������£��ʽ��㣬�ּ��µ�","#00FF00",RndID,-1
				elseif Rnum<60 then 
					RndEvent 1000-Rnum*10,40,rs("��ҵ"),"ҵ��������£����ܴ����ۣ��ּ��µ�","#00FF00",RndID,-1 
				elseif Rnum<70 then
					RndEvent 1000-Rnum*10,30,rs("��ҵ"),"ҵ����ʼ���£��ּ��µ�","#00FF00",RndID,-1
				elseif Rnum<80 then
					RndEvent 1000-Rnum*10,20,rs("��ҵ"),"��˾ҵ�����ѣ�Ͷ�������Ĳ��㣬�ּ��»�","#00FF00",RndID,-1
				elseif Rnum<90 then
					RndEvent 1000-Rnum*10,10,rs("��ҵ"),"�ų���������Ϣ���ּ��»�","#00FF00",RndID,-1
				else
					RndEvent 1000-Rnum*10,1,rs("��ҵ"),"�ų���������Ϣ���ּ�΢��","#00FF00",RndID,-1
				end if			
			end if
		end if '2
		rs.close
	end if '����¼�����
	if cint(AI_Setting(0))=1 then call AI()		'����򿪻������¼�������û�����
	if cint(Gupiao_Setting(12))>0 then
		'ֻ��������Gupiao_Setting(12)���¼�
		conn.execute("delete from RndEvent where id not in(select top "&Gupiao_Setting(12)&" id from RndEvent order by id desc)")
	end if
	
'===========================================================================
' Rate-�۸�仯�Ŀ������ֵ��obj������Flag���ǵ���� -1 ��ʾ�� 1��ʾ��
sub RndEvent(Rmax,Rmin,obj,message,msgColor,GPid,Flag)
	dim content,AMP
	Randomize
	AMP=formatnumber(((Rmax-Rmin+1)*rnd+Rmin)/100,2,true)	
	content="<font color=""#FFFF00"">"&obj&"</font><font color="""&msgColor&""">"&message&" "&AMP&" %</font>"
	conn.execute("insert into RndEvent(content,addtime) values('"&content&"',now())") 
	conn.execute("update [��Ʊ] set ��ǰ�۸�=��ǰ�۸�*"&(1+AMP*Flag/100)&",TodayWave=TodayWave+"&AMP*Flag&",TotalWave=TotalWave+"&AMP*Flag&" where Cid="&GPid&"")
	showmess(content)
end sub
'=============================================================================
sub AI()
	Randomize
	Rnum=rnd
	if Rnum>=1-AI_Setting(1) then
		dim RndID,StockNum,StockName,StockPrice
		dim BuyNum
		RndID=conn.execute("select top 1 cid from ��Ʊ order by cid desc")(0)		'���й�Ʊ����Ŀ  
		Randomize Rnd
		RndID=int(RndID*rnd)+1		'��AI�����Ĺ�ƱID
		set rs=conn.execute("select ������,��ҵ,��ǰ�۸� from ��Ʊ where Cid="&RndID)
		StockNum=rs(0):	StockName=rs(1) : StockPrice=rs(2)
		rs.close
		if StockNum+0<100 then exit sub		'����Ʊ�Ľ�����С��100�ǲ����в���
		'Randomize
		Rnum=int(rnd*10+1)
		select case Rnum
			case 1		'AI�����Ʊ
				BuyNum=int(rnd*10+1)
				call AI_Buy(BuyNum,"������",-1,RndID,StockName,StockPrice,StockNum)
			case 2
				BuyNum=int(rnd*100+1)
				call AI_Buy(BuyNum,"������",-1,RndID,StockName,StockPrice,StockNum)			
			case 4
				BuyNum=int(rnd*50+1)
				call AI_Buy(BuyNum,"������",-1,RndID,StockName,StockPrice,StockNum)			
			case 7
				BuyNum=int(rnd*65+1)
				call AI_Buy(BuyNum,"������",-1,RndID,StockName,StockPrice,StockNum)			
			case 9
				BuyNum=int(rnd*5+1)
				call AI_Buy(BuyNum,"������",-1,RndID,StockName,StockPrice,StockNum)			
			case else
				call AI_Sale("������",-1)
		end select
	elseif Rnum<0.10 then
		conn.execute("insert into RndEvent(content,addtime) values('<font color=white>������</font><font color=#33CC99>�ڹ�Ʊ�г�ת��һȦ��һ�ʽ��׶�û������</font>',now())") 
		'showmess("<font color=white>������</font><font color=#33CC99>�ڹ�Ʊ�г�ת��һȦ��һ�ʽ��׶�û������</font>")
	end if
end sub

sub AI_Buy(BuyNum,AIName,Uid,Sid,GupiaoName,NowPrice,StockNum)
	Dim NewPrice,Zd		'��Ʊ�¼۸�,�ǵ���
	set rs=conn.execute("select �ֹ���,����۸� from �� where sid="&sid&" and uid="&Uid&"")
	if rs.eof and rs.bof then	'1		'���û����������Ʊ
		Randomize
		Rnum=Rnd
		if Rnum<=0.4 then
			Rnum=Rnum+0.5
		end if
		
		NewPrice=NowPrice*(1+BuyNum*Rnum/(StockNum*5)) 		'��Ʊ�¼۸�
		if NewPrice<0 then NewPrice=10
		sql="update ��Ʊ set ��ǰ�۸�="&NewPrice&", ������=������-"&BuyNum&",ʣ��ɷ�=ʣ��ɷ�-"&BuyNum&",�ɽ���=�ɽ���+"&BuyNum&",�������=�������+1 where sid="&sid 
		conn.execute sql
		sql="insert into �� (Uid,�ʺ�,sid,����۸�,ƽ���۸�,�ֹ���,��ҵ) values ("&Uid&",'"&AIName&"','"&sid&"','"&NowPrice&"','"&NowPrice&"','"&BuyNum&"','"&GupiaoName&"')"
		conn.execute sql
		sql="update ai set cash=cash-"&BuyNum*NowPrice&"*1.01,fund=fund+"&BuyNum*NowPrice&"*0.99,StockTypeNum=StockTypeNum+1,TodayBuy=TodayBuy+1,ToTalBuy=ToTalBuy+1 where  AID="&Uid
		conn.execute sql
		
		ZD=(NewPrice-NowPrice)/NowPrice
		mess="<font color=white>"&AIName&"</font><font color=#CC66FF>����"&GupiaoName&" "&BuyNum&" �ɣ��ּ����� "&formatpercent(Zd,2,-1)&"</font>"
		sql="insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )"
		conn.execute sql  
		showmess(mess)
	else
		dim MyNum,AgvPrice	'�ֹ�����ƽ���۸�
		MyNum=rs("�ֹ���")
		
		Randomize
		Rnum=Rnd
		if Rnum<=0.4 then
			Rnum=Rnum+0.5
		end if
		NewPrice=NowPrice*(1+BuyNum*Rnum/(StockNum*5)) 		'��Ʊ�¼۸�
		if NewPrice<0 then NewPrice=10

		AgvPrice=(NowPrice*BuyNum+rs("����۸�")*MyNum)/(BuyNum+MyNum)  	'ƽ���۸�
		
		sql="update ��Ʊ set ��ǰ�۸�="&NewPrice&", ������=������-"&BuyNum&",ʣ��ɷ�=ʣ��ɷ�-"&BuyNum&",�ɽ���=�ɽ���+"&BuyNum&",�������=�������+1 where sid="&sid
		conn.execute sql
		sql="update �� set ƽ���۸�="&AgvPrice&",����۸�="&NowPrice&",�ֹ���=�ֹ���+"&BuyNum&" where Uid="&Uid&" and sid="&sid
		conn.execute sql
		sql="update ai set cash=cash-"&BuyNum*NowPrice&"*1.01,fund=fund+"&BuyNum*NowPrice&"*0.99,TodayBuy=TodayBuy+1,ToTalBuy=ToTalBuy+1 where  AID="&Uid
		conn.execute sql
		
		ZD=(NewPrice-NowPrice)/NowPrice
		mess="<font color=white>"&AIName&"</font><font color=#FF00FF>�ٴ�����"&GupiaoName&" "&BuyNum&" �ɣ��ּ����� "&formatpercent(Zd,2,-1)&"</font>"
		sql="insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"')"
		conn.execute sql  
		showmess(mess)
	end if
	rs.close
	conn.execute("update [GupiaoConfig] set TodayBuy=TodayBuy+1,TodayTotal=TodayTotal+"&BuyNum*NowPrice&"")
end sub	

sub AI_Sale(AIName,AIID)
	dim rst,SaleNum,NowPrice,GupiaoName,Sid,NewPrice	'  , ��������  �� ��Ʊ�ּ� �� ��Ʊ���� �� ��ƱID
	set rst=server.createobject("adodb.recordset")
	sql="select d.�ֹ���,g.��ǰ�۸�,g.��ҵ,d.sid,g.������ from [��] d inner join [��Ʊ] g on d.sid=g.sid where d.Uid="&AIID
	'rst.CursorLocation=3
	rst.open sql,conn,1,1
	if rst.eof and rst.bof then exit sub
	Randomize
	rst.absolutePosition=Fix(rst.recordcount*Rnd)+1
	SaleNum=rst(0)	'�ֹ���
	NowPrice=rst(1)
	GupiaoName=rst(2)
	Sid=rst(3)
	if rst(4)<=0 then exit sub
	if rnd<0.3 then		'������ȫ������ĳ����Ʊ
		Randomize
		Rnum=Rnd
		if Rnum<=0.4 then Rnum=Rnum+0.5
		NewPrice=NowPrice*(1-SaleNum*Rnum/(rst(4)*10))
		if NewPrice<0 then NewPrice=10

		conn.execute "update [��Ʊ] set �ɽ���=�ɽ���+"&SaleNum&",��ǰ�۸�="&NewPrice&", ������=������+"&SaleNum&",ʣ��ɷ�=ʣ��ɷ�+"&SaleNum&",��������=��������+1 where sid="&sid
		conn.execute "delete from [��] where Uid="&AIID&" and sid="&sid
		conn.execute "update [ai] set cash=cash+"&NowPrice*SaleNum&"*0.99,fund=fund-"&NowPrice*SaleNum&"*1.01,StockTypeNum=StockTypeNum-1,TodaySale=TodaySale+1,ToTalSale=ToTalSale+1 where AID="&AIID
	
		mess="<font color=white>"&AIName&"</font><font color=#33CCFF>����"&GupiaoName&" "&SaleNum&" �ɣ��ּ��»� "&formatpercent((NowPrice-NewPrice)/NowPrice,2,-1)&"</font>"
		conn.execute "insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )"  
		showmess(mess)
	else
		Randomize
		Rnum=Rnd
		if Rnum<=0.4 then Rnum=Rnum+0.5
		NewPrice=NowPrice*(1-SaleNum*Rnum/(rst(4)*10))		'�������Ʊ�ļ۸�
		if NewPrice<0 then NewPrice=10

		Rnum=Rnd
		if Rnum>0.5 then Rnum=1-Rnum   
		SaleNum=int(SaleNum*Rnum+1)		' ������Ʊ��������1���Լ��ĳֹ�����*0.5  �������	
		conn.execute "update ��Ʊ set �ɽ���=�ɽ���+"&SaleNum&",��ǰ�۸�="&NewPrice&", ������=������+"&SaleNum&",ʣ��ɷ�=ʣ��ɷ�+"&SaleNum&",��������=��������+1 where sid="&sid
		conn.execute "update �� set �ֹ���=�ֹ���-"&SaleNum&" where Uid="&AIID&" and sid="&sid
		conn.execute "update [ai] set cash=cash+"&NowPrice*SaleNum&"*0.99,fund=fund-"&NowPrice*SaleNum&"*1.01,TodaySale=TodaySale+1,ToTalSale=ToTalSale+1 where AID="&AIID
	
		mess="<font color=white>"&AIName&"</font><font color=#33CCFF>����"&GupiaoName&" "&SaleNum&" �ɣ��ּ��»� "&formatpercent((NowPrice-NewPrice)/NowPrice,2,-1)&"</font>"
		conn.execute "insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )"  
		showmess(mess)
	end if
	rst.close
	set rst=nothing
	conn.execute("update [GupiaoConfig] set TodaySale=TodaySale+1,TodayTotal=TodayTotal+"&NowPrice*SaleNum&"")
end sub	

sub showmess(mess)
'��ʾ��Ϣ��ʼ
dim says,act,towhoway,towho,addwordcolor,saycolor,addsays,saystr,mess1
mess1=mess
mess1=replace(mess1,"white","blue")
mess1=replace(mess1,"#FFFF00","blue")
mess1=replace(mess1,"#00FF00","blue")
if instr(mess1,"������")<>0 then Exit Sub

says="<font color=red><b>����Ʊ��Ϣ��</b></font>"&mess1
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & session("sjjh_name") & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session("nowinroom") & ");<"&"/script>"
addmsg saystr
'����
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