<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html><head>
<META http-equiv=Content-Type content=text/html; charset=gb2312>
<meta name=keywords content="���ֽ�����Ʊ�г�">
<title><%=Gupiao_Setting(5)%>-����ֹ�һ����</title>
<!--#include file="css.asp"-->
</head>
<body topmargin=5 leftmargin=0 oncontextmenu=self.event.returnValue=false>
<%
	dim Uid,UserName,UserCash,Userfund
	dim StockType,TodayBuy,TodaySale,TotalBuy,TotalSale
	Uid=request("uid")
	if not isnumeric(Uid) then
		errmess="<li>����Ĳ�������ָ����ȷ���û�ID"
		call endinfo(2)
	else
		set rs=conn.execute("select * from [�ͻ�] where id="&Uid)
		if rs.eof and rs.bof then
			errmess="<li>ָ�����û�û���ڹ��п���"
			errmess=errmess+"<br><li>ָ�����û������ڣ���ȷ�ϴ��ݵĲ����Ƿ���ȷ"
			call endinfo(2)
		else
			UserName=rs("�ʺ�")
			UserCash=rs("�ʽ�")				
			Userfund=rs("���ʽ�")
			StockType=rs("�ֹ�����")
			TodayBuy=rs("��������")
			TodaySale=rs("��������")
			TotalBuy=rs("������")
			TotalSale=rs("������")
			call main()
		end if
	end if
	CloseDatabase		'�ر����ݿ�
'=====================================	
sub main()		
%>
<table  width="97%" border=0 cellspacing=0 cellpadding=0 align=center bgcolor="#E4E8EF">
	<TR>
		<TD BACKGROUND="Images/topbg.gif" height=9 colspan=3>
	</TD>
	</TR>
	<tr>
		<td valign=center align=middle height=23 background="Images/Header.gif"><font class=big>���� <font color=red><%=UserName%></font> �ֹ�һ��</font></td>
	</tr>
</table>
<br>
<table border="0" width="97%" bgcolor="#666666" cellspacing="1" cellpadding="3" bordercolorlight="#000000" bordercolordark="#6E6E6E"  align="center">
  <TR align=center bgColor=darkblue valign="middle" > 
	  <td background='Images/lan.gif' height=25><font color="#FFFFFF">����</TD>
	  <td background='Images/lan.gif' ><font color="#FFFFFF">��Ʊ����</TD>
	  <td background='Images/lan.gif' ><font color="#FFFFFF">ӵ������</TD>
	  <td background='Images/lan.gif' ><font color="#FFFFFF">�ֹɳɱ�</TD>
	  <td background='Images/lan.gif' ><font color="#FFFFFF">��ǰ�۸�</TD>
	  <td background='Images/lan.gif' ><font color="#FFFFFF">�ֹɼ۸�</TD>
	  <td background='Images/lan.gif' ><font color="#FFFFFF">�ֹ���ֵ</TD>
	  <td background='Images/lan.gif' ><font color="#FFFFFF">ÿ��ӯ��</TD> 
	  <td background='Images/lan.gif' ><font color="#FFFFFF">��ӯ��</TD>
      <td background='Images/lan.gif' ><font color="#FFFFFF">ӯ����</TD> 
  </TR>
<%
	Dim MyGupiaoNum,MyGupiaoValue,MyGupiaoCost 		'���˳ֹ����������˹�Ʊ��ֵ,���˹�Ʊ�ɱ�
	Dim TotalGranRate  '��ӯ����
	MyGupiaoNum=0 :	MyGupiaoValue=0 :	MyGupiaoCost=0:		TotalGranRate=0
	sql="select d.*,g.��ǰ�۸� from [��] d inner join [��Ʊ] g on d.sid=g.sid where d.uid="&Uid&" order by d.[�ֹ���] desc" 
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		RESPONSE.Write "<tr><td colspan=10 bgcolor='#8d9aca'>û�й�Ʊ</tr>"
	else
		Dim rst,CurrentPrice,AgvInPrice		'ĳ��Ʊ��ǰ�۸� ,ĳ��Ʊƽ������۸�
		do while not rs.EOF   
			CurrentPrice=rs("��ǰ�۸�")
			AgvInPrice=iif(rs("ƽ���۸�")=0,rs("����۸�"),rs("ƽ���۸�"))
			MyGupiaoNum=MyGupiaoNum+rs("�ֹ���")
			MyGupiaoValue=MyGupiaoValue+rs("�ֹ���")*CurrentPrice
			MyGupiaoCost=MyGupiaoCost+rs("�ֹ���")*AgvInPrice	
%>
		<TR align="center" bgColor="#8d9aca"> 
			<TD><a href="exchange.asp?sid=<%=rs("sid")%>"><font color="#FFFF00"><%=rs("sid")%></font></a></TD>
			<TD><a href="Dispcompare.asp?sid=<%=rs("sid")%>"><font color="#FFFF00"><%=rs("��ҵ")%></font></a></TD>
			<TD><font color="#FFFF00"><%=rs("�ֹ���")%></font></TD>
			<TD><font color="#FFFF00">��<%=formatnumber(AgvInPrice*rs("�ֹ���"),3,True)%></font></TD>
			<TD><font color="#FFFF00">��<%=formatnumber(CurrentPrice,3,True)%></font></TD>
			<TD><font color="#FFFF00">��<%=formatnumber(AgvInPrice,3,True)%></font></TD>
			<TD><font color="#FFFF00">��<%=formatnumber(CurrentPrice*rs("�ֹ���"),3,True)%></font></TD>
			<TD>
<%
			dim bj,TotalGain,GranRate		'����ӯ��,������Ʊ��ӯ��,������Ʊ��ӯ����
			bj=CurrentPrice-AgvInPrice
			bj=formatnumber(bj,3,True)
			if bj>0 then
				response.write "<font color=#FF0000>��"&bj&" Ԫ</font>"
			elseif bj<0 then
				response.write "<font color=#00FF00>��"&formatnumber(abs(bj),3,True)&" Ԫ</font>"
			elseif bj=0 then
				response.write "<font color=#FFFFff>�� 0.000 Ԫ</font>"
			end if
			response.Write "</TD><TD>"
			
			TotalGain=bj*rs("�ֹ���")
			GranRate=TotalGain*100/MyGupiaoCost
			if TotalGain>0 then 
				response.Write "<font color=#FF0000>��"&formatnumber(TotalGain,3,True)&" Ԫ</font></TD>" 
				response.Write "<TD><font color=#FF0000>��"&formatnumber(GranRate,2,True)&" %</font></TD></tr>"
			elseif TotalGain<0 then
				response.Write "<font color=#00FF00>��"&formatnumber(-TotalGain,3,True)&" Ԫ</font></TD>"
				response.Write "<TD><font color=#00FF00>��"&formatnumber(-GranRate,2,True)&" %</font></TD></tr>" 
			else
				response.Write "<font color=#FFFFff>�� 0.00 Ԫ</font></TD>" 
				response.Write "<TD><font color=#FFFFff>�� 0.00 %</font></TD></tr>"	
			end if	
			rs.MoveNext          
		Loop
		rs.close
	end if 
	if MyGupiaoCost=0 then
		TotalGranRate=0
	else	
		TotalGranRate=(MyGupiaoValue-MyGupiaoCost)*100/MyGupiaoCost
	end if	
%>	
		<TR align=center bgColor="#99CCFF" valign="middle"> 
			<TD background="Images/title.gif" height=20 colspan="2"><font color="#3366CC">�ϣ�������</font></TD> 
			<TD background="Images/title.gif"><font color="#3366CC"><%=formatnumber(MyGupiaoNum,0,-1)%> ��</font></TD>
			<TD background="Images/title.gif"><font color="#3366CC"><%=formatnumber(MyGupiaoCost,2,-1)%> Ԫ</font></TD>
			<td background="Images/title.gif"><font color="#3366CC">-</font></TD>
			<TD background="Images/title.gif"><font color="#3366CC">-</font></TD>
			<TD background="Images/title.gif"><font color="#3366CC"><%=formatnumber(MyGupiaoValue,2,-1)%> Ԫ</font></TD>
			<TD background="Images/title.gif"><font color="#3366CC">-</font></TD>
			<TD background="Images/title.gif"><font color="#3366CC"><%if MyGupiaoValue-MyGupiaoCost>0 then%>+<%end if%><%=formatnumber(MyGupiaoValue-MyGupiaoCost,2,-1)%></font></TD> 
			<TD background="Images/title.gif"><font color="#3366CC"><%if TotalGranRate>0 then%>+<%end if%><%=formatnumber(TotalGranRate,2,-1)%> %</font></TD>
		</TR>
	</TABLE>
<%
%>	
</table>
<br>
<table align=center bgcolor="#0066CC" border=0 cellspacing=1 cellpadding="3" width="97%">
	<TR align=center bgColor=darkblue valign="middle" height="25" > 
		<td background='Images/lan.gif'><B><font color='FFFFFF'>����</font></B></td>
		<td background='Images/lan.gif'><B><font color='FFFFFF'>���ʲ�</font></B></td>
		<td background='Images/lan.gif'><B><font color='FFFFFF'>�ʽ�</font></B></td>
		<td background='Images/lan.gif'><B><font color='FFFFFF'>�ֹ��ܼ�ֵ</font></B></td>
		<td background='Images/lan.gif'><B><font color='FFFFFF'>�ֹ�����</font></B></td>
		<td background='Images/lan.gif'><B><font color='FFFFFF'>��������</font></B></td>
		<td background='Images/lan.gif'><B><font color='FFFFFF'>��������</font></B></td>
		<td background='Images/lan.gif'><B><font color='FFFFFF'>������</font></B></td>
		<td background='Images/lan.gif'><B><font color='FFFFFF'>������</font></B></td></tr> 
	</TR>
	<tr bgcolor='#E4E6EF'align='center' height='22'>
		<td><%=Username%></td>
		<td><%=formatnumber(Userfund,0,true)%> Ԫ</td>
		<td><%=formatnumber(UserCash,0,true)%> Ԫ</td>
		<td><%=formatnumber(MyGupiaoValue,0,true)%> Ԫ</td>
		<td><%=StockType%></td>
		<td><%=TodayBuy%></td>
		<td><%=TodaySale%></td>
		<td><%=TotalBuy%></td>
		<td><%=TotalSale%></td>
	</tr>
	<tr>
		<td align=center colspan=9 background="Images/footer.gif" height=30><a href="TopUser.asp">[�������а�]</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="stock.asp">[���ع���]</a></td>
	</tr>
</table>  
</BODY>
</HTML>
<%
end sub
%>