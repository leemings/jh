<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html><head>
<META http-equiv=Content-Type content=text/html; charset=gb2312>
<meta name=keywords content="���ֽ�����Ʊ�г�">
<title><%=Gupiao_Setting(5)%>-��Ʊ����</title>
<!--#include file="css.asp"-->
</head>
<body oncontextmenu=self.event.returnValue=false; bgcolor="#FFFEF4">
<% 
'=========================================================
' File: Exchange.asp
' Version:1.0
' Date: 2003-1-20
' Script Written by Nev
'=========================================================

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
        '�ָ�Ϊ�����Ĺ�Ʊ
        conn.execute("update ��Ʊ set ��ǰ�۸�=10 where ��ǰ�۸�<0 ")
        conn.execute("update ��Ʊ set ������=0 where ������<0 ")
        
	sql="select ������,��ǰ�۸�,��ҵ,״̬,���̼۸� from [��Ʊ] where sid="&sid
	set rs=conn.execute(sql)
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
		if NowPrice<10 then NowPrice=10
		OpenPrice=rs(4)		'���̼۸� 
		GupiaoName=rs(2)	'��Ʊ���� 
		rs.close
		RaisingLimit=False	'��Ʊ�Ƿ�����ͣ״̬ 
		FallLimit=False		'��Ʊ�Ƿ��ڵ�ͣ״̬ 
			
		if NowPrice>=9999 OR (NowPrice-OpenPrice)/OpenPrice>=Trade_Setting(9)+0 then
			RaisingLimit=True	'��ͣ�� �����벻������
		elseif NowPrice<=Trade_Setting(11)+0 OR (OpenPrice-NowPrice)/OpenPrice>=Trade_Setting(10)+0 then
			FallLimit=True		'��ͣ�� ��������������
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
response.Write "</body></html>"
set rs=nothing
CloseDatabase		'�ر����ݿ�
'======================================
sub main()
	dim sushare,BuyMax
	sql="select �ֹ��� from [��] where uid="&MyUserID&" and sid="&sid
	set rs=conn.execute(sql)
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
	end if 	
	if buymax<0 then buymax=0
	
%>
<table  width="420" height=300 border=0 cellspacing=1 cellpadding=3 align=center bgcolor="#0066CC">
	<TR>
	<TD BACKGROUND="images/topbg.gif" height=9></TD>
	</TR>
	<tr>
		<td valign=center align=middle height=23 class=big background="images/header.gif"><font face="Courier New, Courier, mono" color="#FF0000"><%=sid%></font> �Ź�Ʊ��<%=GupiaoName%></td>
	</tr>
	<tr>
		<td valign=center BGCOLOR=#E4E4EF>
<form name="form1" method="post" action="?action=buy" >
  <input type="hidden" name="sid" value="<%=sid%>">
  <br>
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
	<div align="center"><font color="#006633">����������</font>
		<input name="SaleNum" size="20" value="<%=sushare%>" onFocus="this.select()">
		<input type="submit" value="����" name="B3">
		<input type="reset" value="����" name="B2">
	</div>
	<br><div align="center"><font color="#990000">���Ѿ�ӵ�����ֹ�Ʊ��<font color="#006633"><%=sushare%></font> ��</font></div>
</form>
<div ALIGN=CENTER><FONT color="#3366CC">֤������ѣ����з���Ī�⣬�������У�</FONT></div><br>
		</td>
	</tr>
	<TR><TD height=30 background="images/footer.gif" align="center"><A href="stock.asp" class="cblue">[���д���]</A>&nbsp;&nbsp;<A href="javascript:history.go(-1)" class="cblue">[������һҳ]</a></TD></TR>
</table>
<%
end sub
'-------------------�����Ʊ--------------------------
sub Buy() 
 dim dqjg,sygf,jyl,cjl

	if cint(Trade_Setting(0))=0 then 
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
	if clng(Trade_Setting(4))>0 and MyBiShu>=clng(Trade_Setting(4)) then
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
	elseif BuyNum<clng(Trade_Setting(2)) then
		errmess="<li>�������Ѿ�����ÿ�����ٽ�����Ϊ "&Trade_Setting(2)&" �ɣ�"
		call endinfo(2)
		exit sub	 
	elseif BuyNum>clng(Trade_Setting(3)) then
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
 	        

		set rs=conn.execute("select �ֹ���,ƽ���۸�,BuyTime from [��] where sid="&sid&" and uid="&MyUserID&"")
		
		TotalStockNum=Max((conn.execute("select �ܹɷ� from [��Ʊ] where sid="&sid)(0))/30,1)	'�ܹɷ�
			
		if rs.eof and rs.bof then	'1		'���û����������Ʊ
			Randomize
			Rate=Rnd
			Rnum=Rnd
			if Rnum<=0.4 then Rnum=Rnum+0.5
			if Rate<Trade_Setting(12)+0 then	'����ʱ�۸�����
				NewPrice=BuyNum*Rnum/TotalStockNum 		'��Ʊ�۸�仯��
				ZD=NewPrice/NowPrice
				mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>����"&GupiaoName&" "&BuyNum&" �ɣ��ּ����� "&formatpercent(Zd,2,-1)&"</font>"
			elseif Rate<Trade_Setting(12)+0+Trade_Setting(13) then	'����ʱ�۸��»�
				NewPrice=-1*BuyNum*Rnum/TotalStockNum		'�۸�仯ֵ +
				ZD=NewPrice/NowPrice 
				mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>����"&GupiaoName&" "&BuyNum&" �ɣ��ּ��»� "&formatpercent(-Zd,2,-1)&"</font>"			
			else
				NewPrice=0		'�۸�仯ֵ 0 
				ZD=0
				mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>����"&GupiaoName&" "&BuyNum&" �ɣ��ּ�û�в���</font>"						
			end if
	
			'�жϼ۸�Ŀ����� 
			set rs=conn.execute("select ��ǰ�۸�,ʣ��ɷ�,������,�ɽ��� from [��Ʊ] where sid="&sid)
			if rs.eof then 
			  response.end
			else
			  'dqjg ��ǰ�۸�|sygf ʣ��ɷ�|jyl ������|cjl �ɽ���|
			  dqjg=rs("��ǰ�۸�")+NewPrice
			  if dqjg<10 then dqjg=10  '����10Ԫ����10Ԫ
			  ZD=NewPrice/dqjg
			  if NewPrice<0 then
			    mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>����"&GupiaoName&" "&BuyNum&" �ɣ��ּ��»� "&formatpercent(-Zd,2,-1)&"</font>"			
			  elseif NewPrice=0 then
  			    mess="<font color=#FF6666>"&membername&"</font><font color=#33CCFF>����"&GupiaoName&" "&BuyNum&" �ɣ��ּ�û�в���</font>"						
			  else
			    mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>����"&GupiaoName&" "&BuyNum&" �ɣ��ּ����� "&formatpercent(Zd,2,-1)&"</font>"
			  end if
			  sygf=rs("ʣ��ɷ�")-BuyNum
			  if sygf<0 then sygf=0
			  jyl=rs("������")-BuyNum
			  if jyl<0 then jyl=0
			  cjl=rs("�ɽ���")+BuyNum
			  rs.close
			end if
			
			if ChangeOptB(sid,MyUserID,membername,BuyNum)=true then mess=mess+"<font color=#33CCFF>����ȡ�þ�ӪȨ</font>"
			
			'conn.execute("update [��Ʊ] set ��ǰ�۸�=��ǰ�۸�+"&NewPrice&",ʣ��ɷ�=ʣ��ɷ�-"&BuyNum&",������=������-"&BuyNum&",�ɽ���=�ɽ���+"&BuyNum&",�������=�������+1,TodayWave=TodayWave+"&ZD&",TotalWave=TotalWave+"&ZD&" where sid="&sid)
			conn.execute("update [��Ʊ] set ��ǰ�۸�="&dqjg&",ʣ��ɷ�="&sygf&",������="&jyl&",�ɽ���="&cjl&",�������=�������+1,TodayWave=TodayWave+"&ZD&",TotalWave=TotalWave+"&ZD&" where sid="&sid)
			conn.execute("insert into [��] (Uid,�ʺ�,sid,����۸�,ƽ���۸�,�ֹ���,��ҵ) values ("&MyUserID&",'"&membername&"','"&sid&"','"&NowPrice&"','"&NowPrice&"','"&BuyNum&"','"&GupiaoName&"')")
			conn.execute("update [�ͻ�] set �ʽ�=�ʽ�-"&(DelMoney+Poundage)&",�ֹ�����=�ֹ�����+1,��������=��������+1,������=������+1 where  ID="&MyUserID)
			conn.execute("insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )")  
			showmess(mess)
		elseif datediff("n",rs(2),now())<cint(Trade_Setting(6)) then
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
				NewPrice=BuyNum*Rnum/TotalStockNum		'��Ʊ�۸�仯��
				ZD=NewPrice/NowPrice
				mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>����"&GupiaoName&" "&BuyNum&" �ɣ��ּ����� "&formatpercent(Zd,2,-1)&"</font>"
			elseif Rate<Trade_Setting(12)+0+Trade_Setting(13) then	'����ʱ�۸��»�
				NewPrice=-1*BuyNum*Rnum/TotalStockNum		'�۸�仯ֵ +
				ZD=NewPrice/NowPrice 
				mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>����"&GupiaoName&" "&BuyNum&" �ɣ��ּ��»� "&formatpercent(-Zd,2,-1)&"</font>"			
			else
				NewPrice=0		'�۸�仯ֵ 0 
				ZD=0
				mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>����"&GupiaoName&" "&BuyNum&" �ɣ��ּ�û�в���</font>"						
			end if

			'�жϼ۸�Ŀ����� 
			set rs=conn.execute("select ��ǰ�۸�,ʣ��ɷ�,������,�ɽ��� from [��Ʊ] where sid="&sid)
			if rs.eof then 
			  response.end
			else
			  'dqjg ��ǰ�۸�|sygf ʣ��ɷ�|jyl ������|cjl �ɽ���|
			  dqjg=rs("��ǰ�۸�")+NewPrice
			  if dqjg<10 then dqjg=10  '����10Ԫ����10Ԫ
			  ZD=NewPrice/dqjg
			  if NewPrice<0 then
			    mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>����"&GupiaoName&" "&BuyNum&" �ɣ��ּ��»� "&formatpercent(-Zd,2,-1)&"</font>"			
			  elseif NewPrice=0 then
  			    mess="<font color=#FF6666>"&membername&"</font><font color=#33CCFF>����"&GupiaoName&" "&BuyNum&" �ɣ��ּ�û�в���</font>"						
			  else
			    mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>����"&GupiaoName&" "&BuyNum&" �ɣ��ּ����� "&formatpercent(Zd,2,-1)&"</font>"
			  end if
			  sygf=rs("ʣ��ɷ�")-BuyNum
			  if sygf<0 then sygf=0
			  jyl=rs("������")-BuyNum
			  if jyl<0 then jyl=0
			  cjl=rs("�ɽ���")+BuyNum
			  rs.close
			end if

			if ChangeOptB(sid,MyUserID,membername,BuyNum)=true then mess=mess+"<font color=#33CCFF>����ȡ�þ�ӪȨ</font>"

			AgvPrice=(NowPrice*BuyNum+AgvPrice*MyNum)/(BuyNum+MyNum)  	'��ƽ���۸�			
			'conn.execute("update [��Ʊ] set ��ǰ�۸�=��ǰ�۸�+"&NewPrice&",ʣ��ɷ�=ʣ��ɷ�-"&BuyNum&",������=������-"&BuyNum&",�ɽ���=�ɽ���+"&BuyNum&",�������=�������+1,TodayWave=TodayWave+"&ZD&",TotalWave=TotalWave+"&ZD&" where sid="&sid)
			conn.execute("update [��Ʊ] set ��ǰ�۸�="&dqjg&",ʣ��ɷ�="&sygf&",������="&jyl&",�ɽ���="&cjl&",�������=�������+1,TodayWave=TodayWave+"&ZD&",TotalWave=TotalWave+"&ZD&" where sid="&sid)
			conn.execute("update [��] set ƽ���۸�="&AgvPrice&",����۸�="&NowPrice&",�ֹ���=�ֹ���+"&BuyNum&",BuyTime=now() where Uid="&MyUserID&" and sid="&sid)
			conn.execute("update [�ͻ�] set �ʽ�=�ʽ�-"&(DelMoney+Poundage)&",��������=��������+1,������=������+1 where  ID="&MyUserID)
			conn.execute("insert into RndEvent(content,addtime) values('"&mess&"',now())")  
			showmess(mess)
		end if
		call GivePoundage(sid,Poundage)
		call CalculateFund(MyUserID)
		conn.execute("update [GupiaoConfig] set TodayBuy=TodayBuy+1,TodayTotal=TodayTotal+"&DelMoney&"")
		sucmess="<li>��Ʊ����ɹ���"
		call endinfo(1)
		if cint(Gupiao_Setting(12))>0 then
			'ֻ��������Gupiao_Setting(12)���¼�
			conn.execute("delete from RndEvent where id not in(select top "&Gupiao_Setting(12)&" id from RndEvent order by id desc)")
		end if
		
	end if	'2
end sub
'-------------------������Ʊ--------------------
sub Sale()
     dim dqjg,sygf,jyl,cjl	 

	if cint(Trade_Setting(0))=0 then 
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
	if clng(Trade_Setting(4))>0 and MyBiShu>=clng(Trade_Setting(4)) then
		errmess="<li>�������Ѿ�����һ���������� "&Trade_Setting(4)&" �ʽ��ף����Ѿ��ﵽ��������ޣ�������ٽ��н���"
		call endinfo(1) 
		exit sub	
	end if	
	Dim SaleNum,RemainNum,MyNum,Rnum,Rate,TotalStockNum	'�����������������������Լ���ʣ�������������
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
		
	RemainNum=NowNum+SaleNum
	AddMoney=SaleNum*NowPrice
	Poundage=AddMoney*Trade_Setting(5)/100 	'������
		
	set rs=conn.execute ("select �ֹ���,BuyTime,SaleTime from [��] where Uid="&MyUserID&" and sid="&sid&"")
	
	if rs.eof and rs.bof then 
		errmess="<li>��û�������Ʊ����ô������"
		founderr=true
	elseif datediff("h",rs(1),now())<cint(Trade_Setting(8)) then
		errmess="<li>�������Ѿ���������ĳ��Ʊ����������ʱ����Ϊ��"&Trade_Setting(8)&" Сʱ<br><li>���ϴ�����ù�Ʊ��ʱ���� "&rs(1)&",������׳��ù�Ʊ���"&(Trade_Setting(8)*60-datediff("n",rs(1),now()))&" ����"
		founderr=true				
	elseif datediff("n",rs(2),now())<cint(Trade_Setting(7)) then
		errmess="<li>�������Ѿ�������������ͬһ��Ʊ��ʱ����Ϊ��"&Trade_Setting(7)&" ����<br><li>���ϴ������ù�Ʊ��ʱ���� "&rs(2)&" ����,���´ο����׳��ù�Ʊ���"&(Trade_Setting(7)-datediff("n",rs(2),now()))&" ����"
		founderr=true
	elseif SaleNum<Trade_Setting(2)+0 and rs("�ֹ���")>Trade_Setting(2)+0 then
		errmess="<li>�������Ѿ�����ÿ�����ٽ�����Ϊ "&Trade_Setting(2)&" ��<br><li>������Ĺ�Ʊ������ "&Trade_Setting(2)&" �ɣ�������ȫ������"
		founderr=true
	elseif SaleNum>Trade_Setting(3)+0 then
		errmess="<li>�������Ѿ�����ÿ���������Ϊ "&Trade_Setting(3)&" ֧��Ʊ��"
		founderr=true
	else 			
		MyNum=rs("�ֹ���")-SaleNum		
	end if
	rs.close
	
	if not founderr then
		set rs=conn.execute("select �ܹɷ� from [��Ʊ] where sid="&sid)
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
			ZD=NewPrice/NowPrice 
			mess="<font color=#FF6666>"&membername&"</font><font color=#33CCFF>����"&GupiaoName&" "&SaleNum&" �ɣ��ּ��»� "&formatpercent(-Zd,2,-1)&"</font>"
		elseif Rate<Trade_Setting(14)+0+Trade_Setting(15) then	'����ʱ�۸�����
			NewPrice=NowPrice*SaleNum*Rnum/TotalStockNum		'�۸�仯ֵ +
			ZD=NewPrice/NowPrice 
			mess="<font color=#FF6666>"&membername&"</font><font color=#33CCFF>����"&GupiaoName&" "&SaleNum&" �ɣ��ּ����� "&formatpercent(Zd,2,-1)&"</font>"			
		else
			NewPrice=0		'�۸�仯ֵ 0 
			ZD=0
			mess="<font color=#FF6666>"&membername&"</font><font color=#33CCFF>����"&GupiaoName&" "&SaleNum&" �ɣ��ּ�û�в���</font>"						
		end if			

			'�жϼ۸�Ŀ����� 
			set rs=conn.execute("select ��ǰ�۸�,ʣ��ɷ�,������,�ɽ��� from [��Ʊ] where sid="&sid)
			if rs.eof then 
			  response.end
			else
			  'dqjg ��ǰ�۸�|sygf ʣ��ɷ�|jyl ������|cjl �ɽ���|
			  dqjg=rs("��ǰ�۸�")+NewPrice
			  if dqjg<10 then dqjg=10  '����10Ԫ����10Ԫ
			  ZD=NewPrice/dqjg
			  if NewPrice<0 then
			    mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>����"&GupiaoName&" "&SaleNum&" �ɣ��ּ��»� "&formatpercent(-Zd,2,-1)&"</font>"			
			  elseif NewPrice=0 then
  			    mess="<font color=#FF6666>"&membername&"</font><font color=#33CCFF>����"&GupiaoName&" "&SaleNum&" �ɣ��ּ�û�в���</font>"						
			  else
			    mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>����"&GupiaoName&" "&SaleNum&" �ɣ��ּ����� "&formatpercent(Zd,2,-1)&"</font>"
			  end if
			  sygf=rs("ʣ��ɷ�")+SaleNum
			  cjl=rs("�ɽ���")+SaleNum
			  rs.close
			end if

		
		'conn.execute "update [��Ʊ] set �ɽ���=�ɽ���+"&SaleNum&",ʣ��ɷ�=ʣ��ɷ�+"&SaleNum&",��ǰ�۸�=��ǰ�۸�+"&NewPrice&", ������="&RemainNum&",��������=��������+1,TodayWave=TodayWave+"&ZD&",TotalWave=TotalWave+"&ZD&" where sid="&sid
		conn.execute "update [��Ʊ] set �ɽ���="&cjl&",ʣ��ɷ�="&sygf&",��ǰ�۸�="&dqjg&", ������="&RemainNum&",��������=��������+1,TodayWave=TodayWave+"&ZD&",TotalWave=TotalWave+"&ZD&" where sid="&sid
		conn.execute "delete from [��] where Uid="&MyUserID&" and sid="&sid
		conn.execute "update [�ͻ�] set �ʽ�=�ʽ�+"&AddMoney-Poundage&",�ֹ�����=�ֹ�����-1,��������=��������+1,������=������+1 where ID="&MyUserID

		Rname=ChangeOptS(sid,MyUserID,membername,0)
		if Rname<>"" then mess=mess&"<font color=#33CCFF>,<font color=#FF6666>"&Rname&"</font>��øùɵľ�ӪȨ</font>"
		conn.execute "insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )"  
		showmess(mess)
	else
		Randomize
		Rate=Rnd
		Rnum=Rnd
		if Rnum<0.4 then Rnum=Rnum+0.5
		if Rate<Trade_Setting(14)+0 then	'����ʱ�۸��»�
			NewPrice=-NowPrice*SaleNum*Rnum/TotalStockNum		'�۸�仯ֵ -
			ZD=-NewPrice/NowPrice 
			mess="<font color=#FF6666>"&membername&"</font><font color=#33CCFF>����"&GupiaoName&" "&SaleNum&" �ɣ��ּ��»� "&formatpercent(Zd,2,-1)&"</font>"
		elseif Rate<Trade_Setting(14)+0+Trade_Setting(15) then	'����ʱ�۸�����
			NewPrice=NowPrice*SaleNum*Rnum/TotalStockNum		'�۸�仯ֵ +
			ZD=NewPrice/NowPrice 
			mess="<font color=#FF6666>"&membername&"</font><font color=#33CCFF>����"&GupiaoName&" "&SaleNum&" �ɣ��ּ����� "&formatpercent(Zd,2,-1)&"</font>"			
		else
			NewPrice=0		'�۸�仯ֵ 0 
			ZD=0
			mess="<font color=#FF6666>"&membername&"</font><font color=#33CCFF>����"&GupiaoName&" "&SaleNum&" �ɣ��ּ�û�в���</font>"						
		end if		

			'�жϼ۸�Ŀ����� 
			set rs=conn.execute("select ��ǰ�۸�,ʣ��ɷ�,������,�ɽ��� from [��Ʊ] where sid="&sid)
			if rs.eof then 
			  response.end
			else
			  'dqjg ��ǰ�۸�|sygf ʣ��ɷ�|jyl ������|cjl �ɽ���|
			  dqjg=rs("��ǰ�۸�")+NewPrice
			  if dqjg<10 then dqjg=10  '����10Ԫ����10Ԫ
			  ZD=NewPrice/dqjg
			  if NewPrice<0 then
			    mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>����"&GupiaoName&" "&SaleNum&" �ɣ��ּ��»� "&formatpercent(-Zd,2,-1)&"</font>"			
			  elseif NewPrice=0 then
  			    mess="<font color=#FF6666>"&membername&"</font><font color=#33CCFF>����"&GupiaoName&" "&SaleNum&" �ɣ��ּ�û�в���</font>"						
			  else
			    mess="<font color=#FF6666>"&membername&"</font><font color=#CC66FF>����"&GupiaoName&" "&SaleNum&" �ɣ��ּ����� "&formatpercent(Zd,2,-1)&"</font>"
			  end if
			  sygf=rs("ʣ��ɷ�")+SaleNum
			  cjl=rs("�ɽ���")+SaleNum
			  rs.close
			end if

	
		'conn.execute "update [��Ʊ] set �ɽ���=�ɽ���+"&SaleNum&",ʣ��ɷ�=ʣ��ɷ�+"&SaleNum&",��ǰ�۸�=��ǰ�۸�+"&NewPrice&", ������="&RemainNum&",��������=��������+1,TodayWave=TodayWave+"&ZD&",TotalWave=TotalWave+"&ZD&" where sid="&sid
		conn.execute "update [��Ʊ] set �ɽ���="&cjl&",ʣ��ɷ�="&sygf&",��ǰ�۸�="&dqjg&", ������="&RemainNum&",��������=��������+1,TodayWave=TodayWave+"&ZD&",TotalWave=TotalWave+"&ZD&" where sid="&sid

		conn.execute "update [��] set �ֹ���="&MyNum&",SaleTime=now() where Uid="&MyUserID&" and sid="&sid
		conn.execute "update [�ͻ�] set �ʽ�=�ʽ�+"&AddMoney-Poundage&",��������=��������+1,������=������+1 where ID="&MyUserID
	
		Rname=ChangeOptS(sid,MyUserID,membername,MyNum)
		if Rname<>"" then mess=mess&"<font color=#33CCFF>,<font color=#FF6666>"&Rname&"</font>��øùɵľ�ӪȨ</font>"		
		conn.execute "insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )"  
		showmess(mess)
	end if
	call GivePoundage(sid,Poundage)
	call CalculateFund(MyUserID)
	conn.execute("update [GupiaoConfig] set TodaySale=TodaySale+1,TodayTotal=TodayTotal+"&AddMoney&"")
	sucmess="<li>��Ʊ���������ѳɹ��ύ��"
	call endinfo(1)
	if cint(Gupiao_Setting(12))>0 then
		'ֻ��������Gupiao_Setting(12)���¼�
		conn.execute("delete from RndEvent where id not in(select top "&Gupiao_Setting(12)&" id from RndEvent order by id desc)")
	end if
end sub

'�ж��Ƿ�ȡ�������Ʊ�ľ�ӪȨ
Function ChangeOptB(sid,Uid,UserName,TotalNum)
	ChangeOptB=false
	set rs=conn.execute("select Uid from [��Ʊ] where sid="&sid)
	if rs(0)=0 then		'����ù�Ʊ��û���˾�Ӫ
		conn.execute("update [��Ʊ] set Uid="&Uid&",��Ӫ��='"&UserName&"' where sid="&sid)
		ChangeOptB=true
	elseif rs(0)<>Uid then	'����ù�Ʊ�Ѿ��������˾�Ӫ������Ҫ�ж�˭�Ĺ�Ʊ��
		rs.close  
		set rs=conn.execute("select top 1 �ֹ���,Uid from [��] where sid="&sid&" and Uid<>"&Uid&" and Uid>0 order by �ֹ��� desc")  
		if not rs.eof then
			if TotalNum>rs(0) then 
				conn.execute("update [��Ʊ] set ��Ӫ��='"&UserName&"',Uid="&Uid&" where sid="&sid)
				ChangeOptB=true
			end if	
		end if
		rs.close 	
	end if
end Function

'�ж��Ƿ�ʧȥ�������Ĺ�Ʊ�ľ�ӪȨ,RemNum��������ʣ��ĳֹ���
Function ChangeOptS(sid,Uid,UserName,RemNum)
	ChangeOptS=""
	set rs=conn.execute("select Uid from [��Ʊ] where sid="&sid)
	if rs(0)=Uid then	'������ڵľ�ӪȨ������Ʊ���û�ʱ����Ҫ�ж�
		rs.close
		set rs=conn.execute("select top 1 �ֹ���,�ʺ�,Uid from [��] where sid="&sid&" and Uid>0 and Uid<>"&Uid&" order by �ֹ��� desc")
		if not rs.eof then		'�����Լ�֮�⻹���������������Ʊʱ��Ҫ�����ж�
			if RemNum<rs(0) then    
				conn.execute("update [��Ʊ] set ��Ӫ��='"&rs(1)&"',Uid="&rs(2)&" where sid="&sid) 
				ChangeOptS=rs(1)
			end if
		elseif remnum=0 then		'���ֻ������ӵ�������Ʊ���ұ�������ȫ����Ʊ
			conn.execute("update [��Ʊ] set ��Ӫ��=null,Uid=0 where sid="&sid)
		end if
		rs.close	
	end if	
end Function

'��Ӫ�� ��� ������ 
sub GivePoundage(id,HandleMoney) 
	set rs=conn.execute("select Uid from [��Ʊ] where sid="&id)
	if not rs.eof then
		conn.execute("update [�ͻ�] set �ʽ�=�ʽ�+"&csng(HandleMoney)&",���ʽ�=���ʽ�+"&csng(HandleMoney)&" where id="&rs(0) )
	end if	
end sub

'����Usid�����ʽ� 2003.02.16 
sub CalculateFund(Usid) 
	dim rst,TotalFund
	TotalFund=0	
	sql="select d.�ֹ���,g.��ǰ�۸� from [��] d inner join [��Ʊ] g on d.sid=g.sid where d.uid="&Usid
	set rst=conn.execute(sql)
	do while not rst.eof
		TotalFund=TotalFund+rst(0)*rst(1)
		rst.movenext
	loop
	rst.close
	conn.execute("update [�ͻ�] set ���ʽ�=�ʽ�+"&TotalFund&" where id="&Usid)
end sub

sub showmess(mess)
'��ʾ��Ϣ��ʼ
dim says,addwordcolor,j,username,talkpoint,talkarr,saycolor,mess1
mess1=mess
mess1=replace(mess1,"white","blue")
mess1=replace(mess1,"#FFFF00","blue")
mess1=replace(mess1,"#00FF00","blue")
if instr(mess1,"������")<>0 then Exit Sub

says="<font color=red><b>����Ʊ��Ϣ��</b></font>"&mess1
dim newtalkarr(600) 
   talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=username 
		newtalkarr(595)="���" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)=""&says&""
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
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