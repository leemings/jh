<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html><head>
<META http-equiv=Content-Type content=text/html; charset=gb2312>
<meta name=keywords content="���ֽ�����Ʊ�г�">
<title><%=Gupiao_Setting(5)%>-���й�˾��Ѷ</title>
<!--#include file="css.asp"-->
</head>
<body topmargin=5 leftmargin=0 oncontextmenu=self.event.returnValue=false>
<%
	dim Sid,rsu,TongJi,TongJiF,TotalNum,UserNum
	dim GPScale(),Tnum
	const VHeight=200
	const VWidth=150
	stats="���й�˾��Ѷ"
	Sid=request("Sid")
	if not isnumeric(Sid) then
		errmess="<li>����Ĳ�������ָ����ȷ�Ĺ�Ʊ"
		call endinfo(1)
	else
		set rs=conn.execute("select * from [��Ʊ] where sid="&Sid)
		if rs.eof and rs.bof then
			errmess="<li>û���ҵ�ָ�������й�˾,���ܸ����й�˾�Ѿ�����"
			call endinfo(1)			
		else
			TongJi=rs("TongJi")
			if TongJi="" or isnull(TongJi) then TongJi="0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"
			TongJiF=TongJi
			TongJi=split(TongJi,"|")
			TongJiF=split(TongJiF,"|")
			
			TotalNum=conn.execute("select sum(�ֹ���) from [��] where sid="&sid&" and uid>0")(0)
			UserNum=conn.execute("select count(*) from [��] where sid="&sid&" and uid>0")(0)
			if TotalNum=0 then TotalNum=1
			
			call FormatValue(VHeight)
			Tnum=StockScale(VWidth,UserNum,TotalNum)			
%>
		<table  width="97%" border=0 cellspacing=0 cellpadding=0 align=center bgcolor="#E4E8EF">
			<TR>
				<TD BACKGROUND="Images/topbg.gif" height=9 colspan=3>
			</TD>
			</TR>
			<tr>
				<td valign=center align=middle height=23 background="Images/Header.gif"><font class=big>���й�˾ <font color=red><%=rs("��ҵ")%></font> ��Ѷ</font></td>
			</tr>
		</table>
		<br>
		<table width="97%" border="0" cellspacing="1" cellpadding="3" align="center" style="border: 1px #6595D6 solid; background-color: #FFFFFF;">
			<TR valign="middle"> 
				<td class=forumRow bgcolor="#E4E8EF" align="center" width="150" rowspan="8"><%if rs("LtdImg")<>"" and (not isnull(rs("LtdImg"))) then%><a href="Exchange.asp?sid=<%=rs("sid")%>" title="���뽻��ҳ��"><img border="0" src="<%=rs("LtdImg")%>" width="150"></a><%else%><font color=gray>��ͼƬ</font><%end if%></td>
				<td class=forumRow height="20"><b>���й�˾</b></td>
				<td class=forumRow>&nbsp;<a href="Exchange.asp?sid=<%=rs("sid")%>" class="cblue" title="���뽻��ҳ��"><%=rs("��ҵ")%></a></td>
				<td class=forumRow height="20"><b>��Ӫ��</b></td>
				<td class=forumRow>&nbsp;<%=rs("��Ӫ��")%></td>				
			</TR>
			<TR valign="middle"> 
				<td class=forumRow height="20"><b>��Ʊ����</b></td>
				<td class=forumRow>&nbsp;<%=rs("sid")%></td>
				<td class=forumRow height="20"><b>��������</b></td>
				<td class=forumRow>&nbsp;<%=rs("����")%></td>				
			</TR>			
			<TR valign="middle"> 
				<td class=forumRow height="20"><b>�ܹɷ�</b></td>
				<td class=forumRow>&nbsp;<%=rs("�ܹɷ�")%> ��</td> 
				<td class=forumRow height="20"><b>ʣ��ɷ�</b></td> 
				<td class=forumRow>&nbsp;<%=rs("ʣ��ɷ�")%> ��</td>
			</TR>
			<TR valign="middle"> 
				<td class=forumRow height="20"><b>������</b></td>
				<td class=forumRow>&nbsp;<%=rs("������")%> ��</td> 
				<td class=forumRow height="20"><b>�ɽ���</b></td> 
				<td class=forumRow>&nbsp;<%=rs("�ɽ���")%> ��</td>
			</TR>
			<TR valign="middle"> 
				<td class=forumRow height="20"><b>�������</b></td>
				<td class=forumRow>&nbsp;<%=rs("�������")%> ��</td> 
				<td class=forumRow height="20"><b>��������</b></td> 
				<td class=forumRow>&nbsp;<%=rs("��������")%> ��</td> 
			</TR>
			<TR valign="middle"> 
				<td class=forumRow height="20"><b>���̼۸�</b></td>
				<td class=forumRow>&nbsp;<%=formatcurrency(rs("���̼۸�"),2,-1)%></td>
				<td class=forumRow height="20"><b>��ǰ�۸�</b></td>
				<td class=forumRow>&nbsp;<%=formatcurrency(rs("��ǰ�۸�"),2,-1)%></td>				
			</TR>			
			<TR valign="middle"> 
				<td class=forumRow height="20"><b>�����ǵ�</b></td>
				<td class=forumRow>&nbsp;<%=formatcurrency(rs("��ǰ�۸�")-rs("���̼۸�"),2,true)%></td>  
				<td class=forumRow height="20"><b>�����ǵ���</b></td>   
				<td class=forumRow>&nbsp;<%=formatnumber( (rs("��ǰ�۸�")-rs("���̼۸�"))*100/rs("���̼۸�"),2,true)%> %</td>   
			</TR>	 		
			<TR valign="middle"> 
				<td class=forumRow height="20"><b>����ָ��</b></td>
				<td class=forumRow>&nbsp;<%=formatnumber(rs("TodayWave"),2,-1)%> ��</td> 
				<td class=forumRow height="20"><b>�ۺ�ָ��</b></td>  
				<td class=forumRow>&nbsp;<%=formatnumber(rs("TotalWave"),2,-1)%> ��</td>  
			</TR>									
			<TR valign="middle"> 
				<td class=forumRow height="*" align="center"><b>��˾���</b></td>
				<td class=forumRow colspan="4"><%=rs("Explain")%></td>
			</TR>
		</table> 
		<br>
		<table cellspacing=1 cellpadding=0 width="97%" align=center border=0>
		<tr> 
			<td width="55%"> 	
				<table  width="100%" border=0 cellspacing=1 cellpadding=3 align=center style="border: 1px #6595D6 solid; background-color: #FFFFFF;">													
					<tr>
						<th height=25 align="center" colspan=9 class="admin">���30������ͼ</th>
					</tr> 
					<tr> 
						
						<td class="forumRow" valign="bottom" height="<%=VHeight+20%>" nowrap>
						<%
						for i=0 to ubound(tongji)%>
							<img src="images/bar.gif" width=7 height="<%=TongJiF(i)+5%>" alt="�������̼ۣ�<%=TongJi(i)%> Ԫ">
						<%next%>				
						</td>
					</tr>
				</table>
			</td>
			<td width="3"></td>
			<td>
				<table  width="100%" border=0 cellspacing=1 cellpadding=3 align=center style="border: 1px #6595D6 solid; background-color: #FFFFFF;">													
					<tr>
						<th height=25 align="center" colspan=9 class="admin">ʮ��ɶ��ֹɱ���ͼ</th>
					</tr> 
					<tr><td class="forumRow" valign="top" height="<%=VHeight+20%>">
					<table  width="100%" height='100%' border=0 cellspacing=0 cellpadding=3 align=center>
						<%if UserNum<=0 then
							response.Write "<tr valign=middle align=center><td><font color=gray>��ʱû���˹���ù�Ʊ</td></tr>"
						else
						for i=0 to Tnum%>					
						<tr>
							<td width="90" nowrap><%=GPScale(i,0)%></td>
							<td nowrap><img src="images/baru.gif" height=10 width="<%=GPScale(i,3)+5%>" alt="�ֹ�����<%=GPScale(i,1)%>�� �ֹɰٷֱȣ�<%=formatnumber(GPScale(i,2)*100,2,-1)%>%"></td>
						</tr>
						<%next
						end if
						%>
					</table>
					</td></tr>
				</table>			
			</td>
		</tr>
		</table>				 
		<br>	
		<table  width="97%" border=0 cellspacing=1 cellpadding=3 align=center style="border: 1px #6595D6 solid; background-color: #FFFFFF;">													

		  <TR align=center valign="middle" > 
			  <th height=25>�ɶ�</TD>
			  <th>�ֹ�����</TD>
			  <th>�ֹɱ���</TD>
			  <th>��ǰ�۸�</TD>
			  <th>�ֹɼ۸�</TD>
			  <th>�ֹ���ֵ</TD>  
			  <th>�ֹɳɱ�</TD> 			  
			  <th>��ӯ��</TD>   
			  <th>ӯ����</TD>   
		  </TR> 			
<%
		set rsu=conn.execute("select * from [��] where sid="&sid&" and uid>0 order by �ֹ��� desc") 
		if rsu.eof and rsu.bof then 
			response.Write "<tr><td colspan=9><font color=gray>��ʱû���˹��������Ʊ</td></tr>"
		else
			dim GupiaoValue,GupiaoCost
			do while not rsu.eof 
				GupiaoValue=rs("��ǰ�۸�")*rsu("�ֹ���")		'�ֹ���ֵ
				GupiaoCost =rsu("ƽ���۸�")*rsu("�ֹ���")		'�ֹɳɱ�
%>
		  <TR align=center valign="middle" > 
			  <td class="bg1" height=25><a href='dispu.asp?Uid=<%=rsu("uid")%>' <%if membername=rsu("�ʺ�") then%> class="cblue" <%end if%>><%=rsu("�ʺ�")%></a></TD>
			  <td class="bg1"><%=rsu("�ֹ���")%></TD>
			  <td class="bg1"><%=formatnumber(rsu("�ֹ���")*100/TotalNum,2,-1)%> %</TD>
			  <td class="bg1"><%=formatcurrency(rs("��ǰ�۸�"),2,-1)%></TD>
			  <td class="bg1"><%=formatcurrency(rsu("ƽ���۸�"),2,-1)%></TD>
			  <td class="bg1"><%=formatcurrency(GupiaoValue,2,-1)%></TD>  
			  <td class="bg1"><%=formatcurrency(GupiaoCost,2,-1)%></TD>
			  <td class="bg1"><%=formatcurrency(GupiaoValue-GupiaoCost,2,-1)%></TD>   
			  <td class="bg1"><%if GupiaoCost=0 then%>0.00<%else%><%=formatnumber((GupiaoValue-GupiaoCost)*100/GupiaoCost,2,-1)%><%end if%> %</TD>  
		  </TR>
<%				
				rsu.movenext
			loop
		end if
%>			
		</table>		
		<br>	
		<table  width="97%" border=0 cellspacing=0 cellpadding=0 align=center bgcolor="#E4E8EF">													
			<tr>
				<td align=center background="Images/footer.gif" height=30><a href="Exchange.asp?sid=<%=rs("sid")%>" class="cblue" title="���뽻��">[����]</a>&nbsp;&nbsp;<a href="stock.asp">[���ع���]</a></td>
			</tr>
		</table>  
<%
		end if
		rs.close
	end if
	call GPOnline()	 	'���߹���
	CloseDatabase		'�ر����ݿ�
%>
</BODY>
</HTML>
<%
function FormatValue(StdValue)
	dim Max,ii
	'�ҳ����ֵ
	Max=TongJi(0)
	for ii=1 to ubound(TongJi)
		if csng(Max)<csng(TongJi(ii)) then Max=TongJi(ii)	
	next
	if max=0 then max=1
	'��ʽ������   
	for ii=0 to ubound(TongJi)
	 	TongJiF(ii)=int(TongJi(ii)*StdValue/Max)
	next  
end function   

'10��ɶ��ֹɱ���,StdWidth-��׼���ȣ�TopNum<10 ,TotalN-���йɶ��ֹ���Ŀ
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
	set rst=conn.execute("select top "&num&" �ʺ�,�ֹ��� from [��] where sid="&sid&" and uid>0 order by �ֹ��� desc") 
	for ii=0 to num-1
		GPScale(ii,0)=rst(0)				'�û���
		GPScale(ii,1)=rst(1)				'�ֹ���
		GPScale(ii,2)=rst(1)/TotalN			'�ֹɰٷֱ�	 
		GPScale(ii,3)=int(GPScale(ii,2)*StdWidth)	'��ʽ������  
		StockScale=StockScale+GPScale(ii,1)		'ǰ9���ɶ��ֹ���֮�� 
		rst.movenext 
	next 
	if other then   
		GPScale(9,0)="<font color=gray>�����ɶ�</font>"    
		GPScale(9,1)=TotalN-StockScale
		GPScale(9,2)=GPScale(9,1)/TotalN
		GPScale(9,3)=int(GPScale(ii,2)*StdWidth)
	end if
	StockScale=iif(Topnum>10,9,Topnum-1)
end function
%>