<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_gp_conn.asp"-->
<!--#include file="z_gp_Const.asp"-->
<%
stats="��Ʊ����"
call nav()
call head_var(0,0,GuPiao_Setting(5),"z_gp_gupiao.asp")
response.write "<table class=tableborder1 cellspacing=1 cellpadding=3 align=center border=0 width=" & Forum_body(12) & ">"
if not HaveAccount then
	errmess="<li>����û�й�Ʊ�ʻ������ܳ��ɣ����ȿ���"
	call endinfo(1)
else
	dim sid,mess,BuyTime,SaleTime
	if Request.Form("sid")<>"" then
		sid=trim(Request.Form("sid"))
	else	
		sid=trim(Request.QueryString("sid"))
	end if
	if sid="" or (not isnumeric(sid)) then
		errmess="<li>�Ƿ����ɣ������Ч���ӽ��룡"
		call endinfo(1)
	else
		sql="select JiaoYiLiang,DangQianJiaGe,QiYe,ZhuangTai,KaiPanJiaGe from [GuPiao] where sid="&sid
		set rs=gp_conn.execute(sql)
		if rs.eof and rs.bof then
			rs.close
			errmess="<li>û�������Ʊ"
			call endinfo(1)
		elseif rs(3)="��" then
			rs.close
			errmess="<li>�ù�Ʊ�Ѿ����̣����ܽ��н���"
			call endinfo(2)
		else
			dim NowPrice,NowNum,GupiaoName,OpenPrice,RaisingLimit,FallLimit
			NowNum=rs(0)		'������
			NowPrice=rs(1)		'��ǰ�۸�
			OpenPrice=rs(4)		'���̼۸� 
			GupiaoName=rs(2)	'��Ʊ���� 
			rs.close
			RaisingLimit=False	'��Ʊ�Ƿ�����ͣ״̬ 
			FallLimit=False		'��Ʊ�Ƿ��ڵ�ͣ״̬ 
			if NowPrice>=9999 OR (NowPrice-OpenPrice)/OpenPrice>=Trade_Setting(9)+0 then
				RaisingLimit=True	'��ͣ�� ��������������
			elseif NowPrice<=Trade_Setting(11)+0 OR (OpenPrice-NowPrice)/OpenPrice>=Trade_Setting(10)+0 then
				FallLimit=True		'��ͣ�� �����벻������
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
	'���������Թ���ù�Ʊ���ٹ�
	if NowNum<(int(MyCash*(1-Trade_Setting(5)/100)/nowprice)) and MyCash>=NowNum*nowprice*(1+Trade_Setting(5)/100) then   
		BuyMax=NowNum
	else
		BuyMax=int(MyCash*(1-Trade_Setting(5)/100)/nowprice) 
	end if%>
	<tr>
		<th valign=center align=middle height=23>�� <font color="<%=forum_body(8)%>"><%=sid%></font> �Ź�Ʊ��<%=GupiaoName%></th>
	</tr>
	<tr>
		<td class=tablebody1 valign=center>
			<form name="form1" method="post" action="?action=buy" >
			<input type="hidden" name="sid" value="<%=sid%>"><br>
		  <div align="center"> 
      	<p><font color="#006633">��Ʊ������<font color="#ff0000"><%=NowNum%></font> �ɣ�����ǰ�۸�<font color="#ff0000"><%=formatnumber(nowprice,3)%></font> Ԫ</font></p>
  		</div>
  		<div align="center"> 
      	<p><font color="#3366CC">���Ĺ�Ʊ�ʻ��������ʽ� <font color="#9900CC"><%=formatnumber(MyCash,0)%> </font>Ԫ</font></p>
  		</div>
  		<div align="center"> 
	      <p><font color="#006633">����������</font> 
        <input name="BuyNum" size="20" value="<%=BuyMax%>" onFocus="this.select()">
        <input type="submit" value="����" name="B1">
        <input  type="reset" value="����" name="B2">
        <br><br><font color="#990000">�۳� <%=Trade_Setting(5)%>% �������Ѻ��������Թ��� <font color="#FF0000"><%=BuyMax%> </font>��</font></p>
  		</div>
			</form>
			<form name="form2" method="post" action="?action=sale">
			<input type="hidden" name="sid" value="<%=sid%>">
			<div align="center">
				<font color="#006633">����������</font>
				<input name="SaleNum" size="20" value="<%=sushare%>" onFocus="this.select()">
				<input type="submit" value="����" name="B3">
				<input type="reset" value="����" name="B2">
			</div>
			<br>
			<div align="center">
				<font color="#990000">���Ѿ�ӵ�����ֹ�Ʊ��<font color="#006633"><%=sushare%></font> ��</font>
			</div>
			</form>
			<div ALIGN=CENTER>
				<FONT color="#3366CC">֤������ѣ����з���Ī�⣬�������У�</FONT>
			</div>
			<br>
		</td>
	</tr>
	<TR>
		<Td class=tablebody2 height=25 align="center"><A href="z_gp_Gupiao.asp" class="cblue"><b>[���д���]</b></A>&nbsp;&nbsp;<A href="javascript:history.go(-1)" class="cblue"><b>[������һҳ]</b></a></Td>
	</TR>
<%end sub
'-------------------�����Ʊ--------------------------
sub Buy() 
	if Trade_Setting(0)=0 then 
		errmess="<li>���н����Ѿ���ͣ�����ܽ��й�Ʊ��������"
		call endinfo(1) 
		exit sub	
	end if 
	if RaisingLimit and cint(Trade_Setting(16))<>3 then
		if cint(Trade_Setting(16))<>0 then
			errmess="<li>�������Ѿ����ƶ�����ͣ�Ĺ�Ʊ�ǲ��ܽ����������"
			call endinfo(1) 
			exit sub
		end if
	end if
	if FallLimit and cint(Trade_Setting(17))<>3 then
		if cint(Trade_Setting(17))<>0 then
			errmess="<li>�������Ѿ����ƶ��ڵ�ͣ�Ĺ�Ʊ�ǲ��ܽ����������"
			call endinfo(1) 
			exit sub
		end if
	end if
	if not master and clng(Trade_Setting(4))>0 and MyBiShu>=clng(Trade_Setting(4)) then
		errmess="<li>�������Ѿ�����һ���������� "&Trade_Setting(4)&" �ʽ��ף����Ѿ��ﵽ��������ޣ�������ٽ��н���"
		call endinfo(1) 
		exit sub	
	end if
	Dim BuyNum	'��������
	if Request.Form("BuyNum")="" or (not isnumeric(Request.Form("BuyNum"))) then
		errmess="<li>��Ҫ����ٹ�Ʊ����"
		call endinfo(2)
		exit sub
	end if
	BuyNum=int(Request.Form("BuyNum"))
	if BuyNum<=0 then
		errmess="<li>����������С�ڵ��� 0 �ɣ�"
		call endinfo(2)
		exit sub	
	elseif not master and BuyNum<clng(Trade_Setting(2)) then
		errmess="<li>�������Ѿ�����ÿ�����ٽ�����Ϊ "&Trade_Setting(2)&" �ɣ�"
		call endinfo(2)
		exit sub	 
	elseif not master and BuyNum>clng(Trade_Setting(3)) then
		errmess="<li>�������Ѿ�����ÿ���������Ϊ "&Trade_Setting(3)&" �ɣ�"
		call endinfo(2)
		exit sub
	end if

	dim DelMoney,Poundage	'���Ʊ��Ҫ��Ǯ,������
	DelMoney=BuyNum*NowPrice
	Poundage=DelMoney*Trade_Setting(5)/100		'���� Trade_Setting(5) ��������
	
	if MyCash<DelMoney-Poundage then	'2 
		errmess="<li>�����ʽ����Թ���ָ���Ĺ�Ʊ�������Լ��ٹ����������߹���������Ʊ"
		call endinfo(1)
	elseif int(NowNum)<int(BuyNum) then
		errmess="<li>û���㹻�Ĺ�Ʊ�����������Լ��ٹ����������߹���������Ʊ"
		call endinfo(2)
	else
		Dim Rnum,NewPrice,Zd,Rate,TotalStockNum		'�����,��Ʊ�۸�仯��,�ǵ���
		set rs=gp_conn.execute("select ChiGuShu,PingJunJiaGe,BuyTime from [DaHu] where sid="&sid&" and uid="&MyUserID&"")
		
		TotalStockNum=Max((gp_conn.execute("select ZongGuFen from [GuPiao] where sid="&sid)(0))/30,1)	'�ܹɷ�
			
		if rs.eof and rs.bof then	'1		'���û����������Ʊ
			Randomize
			Rate=Rnd
			Rnum=Rnd
			if Rnum<=0.4 then Rnum=Rnum+0.5
			if Rate<Trade_Setting(12)+0 then	'����ʱ�۸�����
				NewPrice=NowPrice*BuyNum*Rnum/TotalStockNum 		'��Ʊ�۸�仯��
				if NowPrice+NewPrice>OpenPrice*(1+Trade_Setting(9)) then
					NewPrice=OpenPrice*(1+Trade_Setting(9))-NowPrice
				elseif NowPrice+NewPrice<OpenPrice*(1-Trade_Setting(10)) then
					NewPrice=OpenPrice*(1-Trade_Setting(10))-NowPrice
				end if
				ZD=NewPrice/NowPrice
				mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>����"&GupiaoName&" "&BuyNum&" �ɣ��ּ����� "&formatpercent(Zd,2,-1)&"</font>"
			elseif Rate<Trade_Setting(12)+0+Trade_Setting(13) then	'����ʱ�۸��»�
				NewPrice=-NowPrice*BuyNum*Rnum/TotalStockNum		'�۸�仯ֵ +
				if NowPrice+NewPrice>OpenPrice*(1+Trade_Setting(9)) then
					NewPrice=OpenPrice*(1+Trade_Setting(9))-NowPrice
				elseif NowPrice+NewPrice<OpenPrice*(1-Trade_Setting(10)) then
					NewPrice=OpenPrice*(1-Trade_Setting(10))-NowPrice
				end if
				ZD=NewPrice/NowPrice 
				mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>����"&GupiaoName&" "&BuyNum&" �ɣ��ּ��»� "&formatpercent(-Zd,2,-1)&"</font>"			
			else
				NewPrice=0		'�۸�仯ֵ 0 
				ZD=0
				mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>����"&GupiaoName&" "&BuyNum&" �ɣ��ּ�û�в���</font>"						
			end if
			if ChangeOptB(sid,MyUserID,membername,BuyNum)=true then mess=mess+"<font color=#33CCFF>����ȡ�þ�ӪȨ</font>"
			
			gp_conn.execute("update [GuPiao] set DangQianJiaGe=DangQianJiaGe+"&NewPrice&",ShengYuGuFen=ShengYuGuFen-"&BuyNum&",JiaoYiLiang=JiaoYiLiang-"&BuyNum&",ChengJiaoLiang=ChengJiaoLiang+"&BuyNum&",MaiRuBiShu=MaiRuBiShu+1,TodayWave=TodayWave+"&ZD&",TotalWave=TotalWave+"&ZD&" where sid="&sid)
			gp_conn.execute("insert into [DaHu] (Uid,ZhangHao,sid,MaiRuJiaGe,PingJunJiaGe,ChiGuShu,QiYe) values ("&MyUserID&",'"&membername&"','"&sid&"','"&NowPrice&"','"&NowPrice&"','"&BuyNum&"','"&GupiaoName&"')")
			gp_conn.execute("update [KeHu] set ZiJin=ZiJin-"&(DelMoney+Poundage)&",ChiGuZhongLei=ChiGuZhongLei+1,JinRiMaiRu=JinRiMaiRu+1,ZongMaiRu=ZongMaiRu+1 where  ID="&MyUserID)
			gp_conn.execute("insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )")  
			
		elseif not master and datediff("n",rs(2),now())<Trade_Setting(6) then
			errmess="<li>�������Ѿ�������������ͬһ��Ʊ��ʱ����Ϊ��"&Trade_Setting(6)&" ����<br><li>���ϴ�����ù�Ʊ��ʱ���� "&rs(2)&",���´ο��Թ���ù�Ʊ���"&(Trade_Setting(6)-datediff("n",rs(2),now()))&" ����"
			founderr=true
			call endinfo(1)
			exit sub
			
		else
			dim MyNum,AgvPrice	'�ֹ�����ƽ���۸�
			MyNum=rs(0)			'���й�Ʊ��Ŀ			
			AgvPrice=rs(1)		'ƽ���۸�
			Randomize
			Rate=Rnd
			Rnum=Rnd
			if Rnum<=0.4 then Rnum=Rnum+0.5
			if Rate<Trade_Setting(12)+0 then	'����ʱ�۸�����
				NewPrice=NowPrice*BuyNum*Rnum/TotalStockNum		'��Ʊ�۸�仯��
				if NowPrice+NewPrice>OpenPrice*(1+Trade_Setting(9)) then
					NewPrice=OpenPrice*(1+Trade_Setting(9))-NowPrice
				elseif NowPrice+NewPrice<OpenPrice*(1-Trade_Setting(10)) then
					NewPrice=OpenPrice*(1-Trade_Setting(10))-NowPrice
				end if
				ZD=NewPrice/NowPrice
				mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>����"&GupiaoName&" "&BuyNum&" �ɣ��ּ����� "&formatpercent(Zd,2,-1)&"</font>"
			elseif Rate<Trade_Setting(12)+Trade_Setting(13)+0 then	'����ʱ�۸��»�
				NewPrice=-NowPrice*BuyNum*Rnum/TotalStockNum		'�۸�仯ֵ
				if NowPrice+NewPrice>OpenPrice*(1+Trade_Setting(9)) then
					NewPrice=OpenPrice*(1+Trade_Setting(9))-NowPrice
				elseif NowPrice+NewPrice<OpenPrice*(1-Trade_Setting(10)) then
					NewPrice=OpenPrice*(1-Trade_Setting(10))-NowPrice
				end if
				ZD=NewPrice/NowPrice
				mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>����"&GupiaoName&" "&BuyNum&" �ɣ��ּ��»� "&formatpercent(-Zd,2,-1)&"</font>"			
			else
				NewPrice=0		'�۸�仯ֵ 0 
				ZD=0
				mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>����"&GupiaoName&" "&BuyNum&" �ɣ��ּ�û�в���</font>"						
			end if
			if ChangeOptB(sid,MyUserID,membername,MyNum+BuyNum)=true then mess=mess+"<font color=#33CCFF>����ȡ�þ�ӪȨ</font>"

			AgvPrice=(NowPrice*BuyNum+AgvPrice*MyNum)/(BuyNum+MyNum)  	'��ƽ���۸�			
			
			gp_conn.execute("update [GuPiao] set DangQianJiaGe=DangQianJiaGe+"&NewPrice&",ShengYuGuFen=ShengYuGuFen-"&BuyNum&",JiaoYiLiang=JiaoYiLiang-"&BuyNum&",ChengJiaoLiang=ChengJiaoLiang+"&BuyNum&",MaiRuBiShu=MaiRuBiShu+1,TodayWave=TodayWave+"&ZD&",TotalWave=TotalWave+"&ZD&" where sid="&sid)
			gp_conn.execute("update [DaHu] set PingJunJiaGe="&AgvPrice&",MaiRuJiaGe="&NowPrice&",ChiGuShu=ChiGuShu+"&BuyNum&",BuyTime=now() where Uid="&MyUserID&" and sid="&sid)
			gp_conn.execute("update [KeHu] set ZiJin=ZiJin-"&(DelMoney+Poundage)&",JinRiMaiRu=JinRiMaiRu+1,ZongMaiRu=ZongMaiRu+1 where  ID="&MyUserID)
			gp_conn.execute("insert into RndEvent(content,addtime) values('"&mess&"',now())")  
		end if
		call GivePoundage(sid,Poundage)
		call CalculateFund(MyUserID)
		gp_conn.execute("update [GupiaoConfig] set TodayBuy=TodayBuy+1,TodayTotal=TodayTotal+"&DelMoney&"")
		sucmess="<li>��Ʊ����ɹ���"
		call endinfo(1)
		if Gupiao_Setting(12)>0 then
			'ֻ��������Gupiao_Setting(12)���¼�
			gp_conn.execute("delete from RndEvent where id not in(select top "&Gupiao_Setting(12)&" id from RndEvent order by id desc)")
		end if
		
	end if	'2
end sub
'-------------------������Ʊ--------------------
sub Sale()
	if Trade_Setting(0)=0 then 
		errmess="<li>���н����Ѿ���ͣ�����ܽ��й�Ʊ��������"
		call endinfo(1) 
		exit sub	
	end if
	if RaisingLimit and cint(Trade_Setting(16))<>3 then
		if cint(Trade_Setting(16))<>1 then
			errmess="<li>�������Ѿ����ƶ�����ͣ�Ĺ�Ʊ�ǲ��ܽ�����������"
			call endinfo(1) 
			exit sub
		end if
	end if
	if FallLimit and cint(Trade_Setting(17))<>3 then
		if cint(Trade_Setting(17))<>1 then
			errmess="<li>�������Ѿ����ƶ��ڵ�ͣ�Ĺ�Ʊ�ǲ��ܽ�����������"
			call endinfo(1) 
			exit sub
		end if
	end if	
	if not master and clng(Trade_Setting(4))>0 and MyBiShu>=clng(Trade_Setting(4)) then
		errmess="<li>�������Ѿ�����һ���������� "&Trade_Setting(4)&" �ʽ��ף����Ѿ��ﵽ��������ޣ�������ٽ��н���"
		call endinfo(1) 
		exit sub	
	end if	
	Dim SaleNum,MyNum,Rnum,Rate,TotalStockNum	'�����������������������Լ���ʣ�������������
	Dim AddMoney,NewPrice,ZD,Rname,Poundage
	 
	if Request.Form ("SaleNum")="" or (not isnumeric(Request.Form ("SaleNum"))) then
		errmess="<li>������������Ʊ������"
		call endinfo(2)
		exit sub
	elseif Request.Form ("SaleNum")<=0 then
		errmess="<li>�����Ĺ�Ʊ��������Ϊ����߸���"
		call endinfo(2)
		exit sub
	end if
	SaleNum=int(Request.Form ("SaleNum"))
		
	AddMoney=SaleNum*NowPrice
	Poundage=AddMoney*Trade_Setting(5)/100 	'������
		
	set rs=gp_conn.execute ("select ChiGuShu,BuyTime,SaleTime from [DaHu] where Uid="&MyUserID&" and sid="&sid&"")
	
	if rs.eof and rs.bof then 
		errmess="<li>��û�������Ʊ����ô������"
		founderr=true
	elseif not master and datediff("n",rs(1),now())\60<Trade_Setting(8) then
		errmess="<li>�������Ѿ���������ĳ��Ʊ����������ʱ����Ϊ��"&Trade_Setting(8)&" Сʱ<br><li>���ϴ�����ù�Ʊ��ʱ���� "&rs(1)&",������׳��ù�Ʊ���"&(Trade_Setting(8)*60-datediff("n",rs(1),now()))&" ����"
		founderr=true				
	elseif not master and datediff("n",rs(2),now())<Trade_Setting(7) then
		errmess="<li>�������Ѿ�������������ͬһ��Ʊ��ʱ����Ϊ��"&Trade_Setting(7)&" ����<br><li>���ϴ������ù�Ʊ��ʱ���� "&rs(2)&" ����,���´ο����׳��ù�Ʊ���"&(Trade_Setting(7)-datediff("n",rs(2),now()))&" ����"
		founderr=true
	elseif not master and SaleNum<Trade_Setting(2)+0 and rs("ChiGuShu")>Trade_Setting(2)+0 then
		errmess="<li>�������Ѿ�����ÿ�����ٽ�����Ϊ "&Trade_Setting(2)&" ��<br><li>������Ĺ�Ʊ������ "&Trade_Setting(2)&" �ɣ�������ȫ������"
		founderr=true
	elseif not master and SaleNum>Trade_Setting(3)+0 then
		errmess="<li>�������Ѿ�����ÿ���������Ϊ "&Trade_Setting(3)&" ֧��Ʊ��"
		founderr=true
	else 			
		MyNum=rs("ChiGuShu")-SaleNum		
	end if
	rs.close
	
	if not founderr then
		set rs=gp_conn.execute("select ZongGuFen from [GuPiao] where sid="&sid)
		if rs.eof and rs.bof then
			errmess="<li>�ù�Ʊ�Ѿ�ɾ�����������Ա��ϵ"
			founderr=true		
		else
			TotalStockNum=Max(rs(0)/30,1)	'�ܹɷ�
		end if
	end if
	
	if founderr then
		call endinfo(1)
		exit sub
	end if	
			
	if MyNum<0 then			'û����ô���Ʊ
		errmess="<li>��û����ô���Ʊ����ô������"
		call endinfo(2)
		exit sub	
	elseif MyNum=0 then		'ȫ������ 
		Randomize
		Rate=Rnd
		Rnum=Rnd
		if Rnum<=0.4 then Rnum=Rnum+0.5
		if Rate<Trade_Setting(14)+0 then	'����ʱ�۸��»�
			NewPrice=-NowPrice*SaleNum*Rnum/TotalStockNum		'�۸�仯ֵ -
			if NowPrice+NewPrice>OpenPrice*(1+Trade_Setting(9)) then
				NewPrice=OpenPrice*(1+Trade_Setting(9))-NowPrice
			elseif NowPrice+NewPrice<OpenPrice*(1-Trade_Setting(10)) then
				NewPrice=OpenPrice*(1-Trade_Setting(10))-NowPrice
			end if
			ZD=NewPrice/NowPrice 
			mess="<font color=#FF6666>"&membername&"</font><font color=#33CCFF>����"&GupiaoName&" "&SaleNum&" �ɣ��ּ��»� "&formatpercent(-Zd,2,-1)&"</font>"
		elseif Rate<Trade_Setting(14)+Trade_Setting(15)+0 then	'����ʱ�۸�����
			NewPrice=NowPrice*SaleNum*Rnum/TotalStockNum		'�۸�仯ֵ +
			if NowPrice+NewPrice>OpenPrice*(1+Trade_Setting(9)) then
				NewPrice=OpenPrice*(1+Trade_Setting(9))-NowPrice
			elseif NowPrice+NewPrice<OpenPrice*(1-Trade_Setting(10)) then
				NewPrice=OpenPrice*(1-Trade_Setting(10))-NowPrice
			end if
			ZD=NewPrice/NowPrice 
			mess="<font color=#FF6666>"&membername&"</font><font color=#33CCFF>����"&GupiaoName&" "&SaleNum&" �ɣ��ּ����� "&formatpercent(Zd,2,-1)&"</font>"			
		else
			NewPrice=0		'�۸�仯ֵ 0 
			ZD=0
			mess="<font color=#FF6666>"&membername&"</font><font color=#33CCFF>����"&GupiaoName&" "&SaleNum&" �ɣ��ּ�û�в���</font>"						
		end if			
		
		gp_conn.execute "update [GuPiao] set ChengJiaoLiang=ChengJiaoLiang+"&SaleNum&",ShengYuGuFen=ShengYuGuFen+"&SaleNum&",DangQianJiaGe=DangQianJiaGe+"&NewPrice&", JiaoYiLiang=JiaoYiLiang+"&SaleNum&",MaiChuBiShu=MaiChuBiShu+1,TodayWave=TodayWave+"&ZD&",TotalWave=TotalWave+"&ZD&" where sid="&sid
		gp_conn.execute "delete from [DaHu] where Uid="&MyUserID&" and sid="&sid
		gp_conn.execute "update [KeHu] set ZiJin=ZiJin+"&AddMoney-Poundage&",ChiGuZhongLei=ChiGuZhongLei-1,JinRiMaiChu=JinRiMaiChu+1,ZongMaiChu=ZongMaiChu+1 where ID="&MyUserID

		Rname=ChangeOptS(sid,MyUserID,membername,0)
		if Rname<>"" then mess=mess&"<font color=#33CCFF>,<font color=#FF6666>"&Rname&"</font>��øùɵľ�ӪȨ</font>"
		gp_conn.execute "insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )"  
	else
		Randomize
		Rate=Rnd
		Rnum=Rnd
		if Rnum<0.4 then Rnum=Rnum+0.5
		if Rate<Trade_Setting(14)+0 then	'����ʱ�۸��»�
			NewPrice=-NowPrice*SaleNum*Rnum/TotalStockNum		'�۸�仯ֵ -
			if NowPrice+NewPrice>OpenPrice*(1+Trade_Setting(9)) then
				NewPrice=OpenPrice*(1+Trade_Setting(9))-NowPrice
			elseif NowPrice+NewPrice<OpenPrice*(1-Trade_Setting(10)) then
				NewPrice=OpenPrice*(1-Trade_Setting(10))-NowPrice
			end if
			ZD=-NewPrice/NowPrice 
			mess="<font color=#FF6666>"&membername&"</font><font color=#33CCFF>����"&GupiaoName&" "&SaleNum&" �ɣ��ּ��»� "&formatpercent(Zd,2,-1)&"</font>"
		elseif Rate<Trade_Setting(14)+Trade_Setting(15)+0 then	'����ʱ�۸�����
			NewPrice=NowPrice*SaleNum*Rnum/TotalStockNum		'�۸�仯ֵ +
			if NowPrice+NewPrice>OpenPrice*(1+Trade_Setting(9)) then
				NewPrice=OpenPrice*(1+Trade_Setting(9))-NowPrice
			elseif NowPrice+NewPrice<OpenPrice*(1-Trade_Setting(10)) then
				NewPrice=OpenPrice*(1-Trade_Setting(10))-NowPrice
			end if
			ZD=NewPrice/NowPrice 
			mess="<font color=#FF6666>"&membername&"</font><font color=#33CCFF>����"&GupiaoName&" "&SaleNum&" �ɣ��ּ����� "&formatpercent(Zd,2,-1)&"</font>"			
		else
			NewPrice=0		'�۸�仯ֵ 0 
			ZD=0
			mess="<font color=#FF6666>"&membername&"</font><font color=#33CCFF>����"&GupiaoName&" "&SaleNum&" �ɣ��ּ�û�в���</font>"						
		end if		
		gp_conn.execute "update [GuPiao] set ChengJiaoLiang=ChengJiaoLiang+"&SaleNum&",ShengYuGuFen=ShengYuGuFen+"&SaleNum&",DangQianJiaGe=DangQianJiaGe+"&NewPrice&", JiaoYiLiang=JiaoYiLiang+"&SaleNum&",MaiChuBiShu=MaiChuBiShu+1,TodayWave=TodayWave+"&ZD&",TotalWave=TotalWave+"&ZD&" where sid="&sid
		gp_conn.execute "update [DaHu] set ChiGuShu="&MyNum&",SaleTime=now() where Uid="&MyUserID&" and sid="&sid
		gp_conn.execute "update [KeHu] set ZiJin=ZiJin+"&AddMoney-Poundage&",JinRiMaiChu=JinRiMaiChu+1,ZongMaiChu=ZongMaiChu+1 where ID="&MyUserID
	
		Rname=ChangeOptS(sid,MyUserID,membername,MyNum)
		if Rname<>"" then mess=mess&"<font color=#33CCFF>,<font color=#FF6666>"&Rname&"</font>��øùɵľ�ӪȨ</font>"		
		gp_conn.execute "insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )"  
	end if
	call GivePoundage(sid,Poundage)
	call CalculateFund(MyUserID)
	gp_conn.execute("update [GupiaoConfig] set TodaySale=TodaySale+1,TodayTotal=TodayTotal+"&AddMoney&"")
	sucmess="<li>��Ʊ���������ѳɹ��ύ��"
	call endinfo(1)
	if Gupiao_Setting(12)>0 then
		'ֻ��������Gupiao_Setting(12)���¼�
		gp_conn.execute("delete from RndEvent where id not in(select top "&Gupiao_Setting(12)&" id from RndEvent order by id desc)")
	end if
end sub

'�ж��Ƿ�ȡ�������Ʊ�ľ�ӪȨ
Function ChangeOptB(sid,Uid,UserName,TotalNum)
	ChangeOptB=false
	set rs=gp_conn.execute("select Uid from [GuPiao] where sid="&sid)
	if rs(0)=0 then		'����ù�Ʊ��û���˾�Ӫ
		gp_conn.execute("update [GuPiao] set Uid="&Uid&",JingYingZhe='"&UserName&"' where sid="&sid)
		ChangeOptB=true
	elseif rs(0)<>Uid then	'����ù�Ʊ�Ѿ��������˾�Ӫ������Ҫ�ж�˭�Ĺ�Ʊ��
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

'�ж��Ƿ�ʧȥ�������Ĺ�Ʊ�ľ�ӪȨ,RemNum��������ʣ��ĳֹ���
Function ChangeOptS(sid,Uid,UserName,RemNum)
	ChangeOptS=""
	set rs=gp_conn.execute("select Uid from [GuPiao] where sid="&sid)
	if rs(0)=Uid then	'������ڵľ�ӪȨ������Ʊ���û�ʱ����Ҫ�ж�
		rs.close
		set rs=gp_conn.execute("select top 1 ChiGuShu,ZhangHao,Uid from [DaHu] where sid="&sid&" and Uid>0 and Uid<>"&Uid&" order by ChiGuShu desc")
		if not rs.eof then		'�����Լ�֮�⻹���������������Ʊʱ��Ҫ�����ж�
			if RemNum<rs(0) then    
				gp_conn.execute("update [GuPiao] set JingYingZhe='"&rs(1)&"',Uid="&rs(2)&" where sid="&sid) 
				ChangeOptS=rs(1)
			end if
		elseif remnum=0 then		'���ֻ������ӵ�������Ʊ���ұ�������ȫ����Ʊ
			gp_conn.execute("update [GuPiao] set JingYingZhe=null,Uid=0 where sid="&sid)
		end if
		rs.close	
	end if	
end Function

'��Ӫ�� ��� ������ 
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