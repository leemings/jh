<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html><head>
<META http-equiv=Content-Type content=text/html; charset=gb2312>
<meta name=keywords content="��E�߽�������Ʊ�г�">
<META HTTP-EQUIV="luxiaoqing" CONTENT="no-cache">
<meta http-equiv="refresh" content="<%=Gupiao_Setting(1)%>">
<title><%=Gupiao_Setting(5)%>-<%=membername%></title>
<!--#include file="css.asp"-->
</head><body bgcolor="#ffffff" text="#000000" style="FONT-SIZE: 9pt" topmargin=5 leftmargin=0 oncontextmenu=self.event.returnValue=false onselectstart=self.event.returnValue=false>
<%
        conn.execute("update ��Ʊ set ������=0 where ������<0 ")
%>
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
		<td valign=center align=middle height=23 background="Images/Header.gif"><font class="big"><b><%=Gupiao_Setting(5)%>����</b></font></td>
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
			<td height="25" background='images/lan.gif' align="center">&nbsp;<font color="#00CC00">�� <a href="announcements.asp"><font color="#FFFFFF">���й����</font></a> ��</font></td>
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
				messages[0]="��ӭ���ٹ�Ʊ������"
				mescolor[0]="blue"
				messages[1]="֤������ѣ����з���Ī�⣬�������У�"
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
			messages[<%=i%>]="֤������ѣ����з���Ī�⣬�������У�"
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
			
          <td height=25 background='images/lan.gif' align=center><font color="#FFFFFF">���ӹ���</font></td>
		</tr>		
		<tr valign=top bgcolor="#000000"> 
			<td align=center width="100%"> 
				<table cellspacing=1 cellpadding=0 border=0 width="198" height="54">
					<tr>
						<td><img src="images/clock/cb.gif" name="a"><img src="images/clock/cb.gif" name="b"><img src="images/clock/colon.gif" name="c"><img src="images/clock/cb.gif" name="d"><img src="images/clock/cb.gif" name="e"><img src="images/clock/colon.gif" name="f"><img src="images/clock/cb.gif" name="g"><img src="images/clock/cb.gif" name="h"><img src="images/clock/cam.gif" name="j"></td>
						<td align="right"><font color="#CCCCCC"><%=WeekdayName(Weekday(Date))%></font>&nbsp;</td>
					</tr>
					<tr>
                <td colspan="2" align="right"><font color="#CCCCCC"><%=DatePart("yyyy", Date) & " �� " & DatePart("m", Date) & " �� " & DatePart("d", Date) & " ��"%></font>&nbsp;</td>
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
    <td height="25" align="center" background="Images/lan.gif"><font color=#ffffff>����</font></td>
	<td align="center" background="Images/lan.gif"><font color=#ffffff>��Ʊ����</font></td>
	<td align="center" background="Images/lan.gif"><font color=#ffffff>��Ӫ��</font></td>
	<td align="center" background="Images/lan.gif"><font color=#ffffff>������</font></td>
 	<td align="center" background="Images/lan.gif"><font color=#ffffff>�����̼�</font></td>
	<td align="center" background="Images/lan.gif"><font color=#ffffff>�ɽ���</font></td>
	<td align="center" background="Images/lan.gif"><font color=#ffffff>�������</font></td>
	<td align="center" background="Images/lan.gif"><font color=#ffffff>��������</font></td>
	<td align="center" background="Images/lan.gif"><font color=#ffffff>�ɽ���</font></td>
	<td align="center" background="Images/lan.gif"><font color=#ffffff>�ǵ���</font></td>
  </tr>
<%  
	dim CurrentPrice,ZangDieRange,sstop	'��Ʊ��ǰ�۸񣬹�Ʊ�ǵ����ȣ���Ʊ״̬(ͣ�ƣ���ͣ����ͣ��)
	Dim ColExp 		'�ۺ�ָ��
	Dim color
	stats="���д���"
	call GPOnline()	 	'���߹���        
	sql= "select * from ��Ʊ order by sid"         
	set rs=conn.execute(sql)         
	DO while not RS.EOF  
		if rs("��ǰ�۸�")<=1 then           
			color="#ffffff"           
		elseif rs("���̼۸�")<rs("��ǰ�۸�") then           
			color="#FF0000"           
		elseif rs("���̼۸�")>rs("��ǰ�۸�") then          
			color="#00FF00"           
		elseif rs("���̼۸�")=rs("��ǰ�۸�") then             
			color="#FFCC00"          
		end if          
		ZangDieRange=csng((rs("��ǰ�۸�")-rs("���̼۸�"))/rs("���̼۸�"))
		CurrentPrice=formatnumber(rs("��ǰ�۸�"),2,true)
		if CurrentPrice<=Trade_Setting(11)+0 then
			sstop=4
			ZangDieRange="<b>ͣ��</b>"
		elseif ZangDieRange>=Trade_Setting(9)+0 then          
			sstop=1
			ZangDieRange="<b>��ͣ</b>"         
		elseif ZangDieRange<=-Trade_Setting(10)+0 then 
			sstop=2 
			ZangDieRange="<b>��ͣ</b>" 
		elseif ZangDieRange=0 then
			sstop=3
			ZangDieRange="0.00 %"
		else       
			sstop=0
		end if   
%>
  <tr bgcolor="#3E3E3E" valign="bottom"> 
    <td height="20" width="6%" align="center"><a href="exchange.asp?sid=<%=rs("sid")%>"><font color=#FFCC00><%=rs("sid")%></font></a></td>
    <td height="20" width="14%" align="center"><a href="DispCompare.asp?sid=<%=rs("sid")%>"><font color="<%=color%>"><%=rs("��ҵ")%></font></a></td>
    <td height="20" align="center"><%=iif(rs("��Ӫ��")="","<font color=#FFffff>-</font>","<a href='dispu.asp?Uid="&rs("uid")&"'><font color=#FFCC00>"&rs("��Ӫ��")&"</font></a>")%></td>
	<td height="20" align="center"><font color=#FFCC00><%=rs("������")%></font></td>
    <td height="20" align="center"><font color=#FFCC00><%=formatcurrency(rs("���̼۸�"),2,true)%></font></td>
    <td height="20" align="center"><font color="<%=color%>"><%=formatcurrency(rs("��ǰ�۸�"),2,true)%></font></td>
    <td height="20" align="center"><%=iif(rs("�������")=0,"<font color=#ffffff>0</font>","<font color=#FFCC00>"&rs("�������")&"</font>")%></td>
	<td height="20" align="center"><%=iif(rs("��������")=0,"<font color=#ffffff>0</font>","<font color=#FFCC00>"&rs("��������")&"</font>")%></td>
	<td height="20" align="right"><%=iif(rs("�ɽ���")=0,"<font color=#ffffff>0&nbsp;</font>","<font color=#FFCC00>"&rs("�ɽ���")&"&nbsp;</font>")%></td>
    <td height="20" align="center"><font color="<%=color%>"> 
<% 
		if sstop<>0 then
			response.write ""&ZangDieRange&""
		else
			response.write ""&formatnumber(ZangDieRange*100,2,true)&" %"
			ColExp=ColExp+ZangDieRange*100		'�����ۺ�ָ��
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
'==============���˹�Ʊ����=============
	if HaveAccount then
%>
	<br>
	<TABLE cellSpacing=1 cellPadding=2 width="98%" align=center bgColor="#666666" border=0>
		<TR bgcolor="#666666"> 
			<TD colSpan=10 height=23 valign="middle" background="Images/header.gif">&nbsp;<font color="#0066FF">�� <font color="#FF0000">���ɹ���</font> ��</font>&nbsp;&nbsp;��ӭ����<b><%=membername%></b>�����ĳֹ�������£�</TD>
		</TR>
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
	Dim rst,AgvInPrice		'ĳ��Ʊ��ǰ�۸� ,ĳ��Ʊƽ������۸�   
	Dim MyGupiaoNum,MyGupiaoValue,MyGupiaoCost 		'���˳ֹ����������˹�Ʊ��ֵ,���˹�Ʊ�ɱ�             
	Dim YKStr,TotalGranRate  'ӯ�����,��ӯ����
	MyGupiaoNum=0 :	MyGupiaoValue=0 :	MyGupiaoCost=0
	sql="select d.*,g.��ǰ�۸� from [��] d inner join [��Ʊ] g on d.sid=g.sid where d.uid="&MyUserID&" order by d.[�ֹ���] desc"  
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		RESPONSE.Write "<tr><td colspan=10 bgcolor='#8d9aca'> ��������û�й�Ʊ</tr>"
	else
		do while not rs.EOF   
			CurrentPrice=rs("��ǰ�۸�")
			AgvInPrice=rs("ƽ���۸�")
			MyGupiaoNum=MyGupiaoNum+rs("�ֹ���")
			MyGupiaoValue=MyGupiaoValue+rs("�ֹ���")*CurrentPrice
			MyGupiaoCost=MyGupiaoCost+rs("�ֹ���")*AgvInPrice	
%>
		<TR align="center" bgColor="#8d9aca"> 
			<TD><a href="exchange.asp?sid=<%=rs("sid")%>"><font color="#FFFF00"><%=rs("sid")%></font></a></TD>
			<TD><a href="exchange.asp?sid=<%=rs("sid")%>"><font color="#FFFF00"><%=rs("��ҵ")%></font></a></TD>
			<TD><font color="#FFFF00"><%=rs("�ֹ���")%></font></TD>
			<TD><font color="#FFFF00">��<%=formatnumber(AgvInPrice*rs("�ֹ���"),2,True)%></font></TD>
			<TD><font color="#FFFF00">��<%=formatnumber(CurrentPrice,2,True)%></font></TD>
			<TD><font color="#FFFF00">��<%=formatnumber(AgvInPrice,2,True)%></font></TD>
			<TD><font color="#FFFF00">��<%=formatnumber(CurrentPrice*rs("�ֹ���"),2,True)%></font></TD>
			<TD>
<%
			dim bj,TotalGain,GranRate		'����ӯ��,������Ʊ��ӯ��,������Ʊ��ӯ����
			bj=CurrentPrice-AgvInPrice
			bj=formatnumber(bj,2,True)
			if bj>0 then
				response.write "<font color=#FF0000>��"&bj&" Ԫ</font>"
			elseif bj<0 then
				response.write "<font color=#00FF00>��"&formatnumber(abs(bj),2,True)&" Ԫ</font>"
			elseif bj=0 then
				response.write "<font color=#FFFFff>�� 0.000 Ԫ</font>"
			end if
			response.Write "</TD><TD>"
			
			TotalGain=bj*rs("�ֹ���")
			GranRate=TotalGain*100/MyGupiaoCost
			if TotalGain>0 then 
				response.Write "<font color=#FF0000>��"&formatnumber(TotalGain,2,True)&" Ԫ</font></TD>" 
				response.Write "<TD><font color=#FF0000>��"&formatnumber(GranRate,2,True)&" %</font></TD></tr>"
			elseif TotalGain<0 then
				response.Write "<font color=#00FF00>��"&formatnumber(-TotalGain,2,True)&" Ԫ</font></TD>"
				response.Write "<TD><font color=#00FF00>��"&formatnumber(-GranRate,2,True)&" %</font></TD></tr>" 
			else
				response.Write "<font color=#FFFFff>�� 0.00 Ԫ</font></TD>" 
				response.Write "<TD><font color=#FFFFff>�� 0.00 %</font></TD></tr>"	
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
	YKStr=YKStr&" Ԫ</font>" 
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
	<td valign=top width="29%"> 		<!--��һ��-->
	
		<table cellspacing=1 cellpadding=0 width="100%" border=0 bgcolor="0066CC">
			<tr bgcolor="#E4E8EF"> 
				<td height=25 background='images/lan.gif' align=center><font color="#FFFFFF">���չ�������</font></td>
			</tr>		
        	<tr valign=top bgcolor="#E4E8EF"> 
          		<td align=center width="100%"> 
					<table cellspacing=1 cellpadding=0 border=0 width="200">
						<tr><td colspan="2" height="3"></td></tr>
						<tr><td>�ɽ��ܶ</td><td align=right><%=ChengJiaoMoney%> Ԫ</td></tr>
						<tr><td>�ɽ�������</td><td align=right><%=ChengJiaoNum%> ��</td></tr>
						<tr><td>�ۺ�ָ����</td><td align=right><%if ColExp>0 then%><font color=#FF0000>��<%elseif ColExp<0 then%><font color="#0066FF">��<%else%><font color="#0066FF">�� <%end if%><%=formatnumber(abs(ColExp),2,true)%>��%</font></td></tr>
						<tr><td colspan="2" height="3"></td></tr>
					</table>			
				</td>
			</tr>
		</table>
        <br>
		<table cellspacing=1 cellpadding=0 border=0 bgcolor="#0066CC" width="100%">
			<tr> 
				<td colspan="2"  height=25 background='images/lan.gif' align=center> <font color="#FFFFFF">�����������</font></td>
			</tr>		
<%
		if HaveAccount then			
%>          
			<tr bgcolor="#E4E8EF"> 
				<td align=center width="100%"> 
					<table cellspacing=1 cellpadding=0 border=0 width="200"> 
						<tr><td colspan="2" height="3"></td></tr>
						<tr>
							<td>�ʻ��ʽ�</td> 
							<td align=right><%=formatnumber(MyCash,2,-1)%> Ԫ</td>
						</tr>
						<tr>
							<td>�ֹ�������</td>  
							<td align=right><%=formatnumber(MyGupiaoNum,0,-1)%> ��</td>
						</tr>						
						<tr>
							<td>�ֹ���ֵ��</td>  
							<td align=right><%=formatnumber(MyGupiaoValue,2,-1)%> Ԫ</td>
						</tr>
						<tr>
							<td>��Ʊ�ɱ���</td>
							<td align=right><%=formatnumber(MyGupiaoCost,2,-1)%> Ԫ</td>
						</tr>
						<tr>
							<td>ӯ�������</td>
							<td align=right><%=YKStr%></td>
						</tr>												
						<tr>
							<td>�ʲ��ϼƣ�</td>  
							<td align=right><%=formatnumber(MyGupiaoValue+MyCash,2,-1)%> Ԫ</td>  
						</tr>
						<tr> 
							<td>����������</td> 
							<td align=right><%=MyBiShu%> ��</td> 
						</tr>
						<tr><td colspan="2" height="3"></td></tr>	 
					</table>
				</td>
			</tr>			
			<tr>
				<td align=center background="images/title.gif" height=21 bgcolor="#E4E8EF"><a href="Mymanage.asp" class="cblue">�˻�����</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="guide/guide1.asp" class="cred">����ָ��</a></td>
			</tr>
					
		<% else %>
			<form name="form1" method="post" action="Mymanage.asp?action=new">
			<tr>
				<td width=250 bgcolor="#ffffff"> 
					<input type=hidden name=username value="<%=session("sjjh_name")%>">
					<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color="#9933CC">���ڹ�Ʊ�г��ﻹû���ʻ���Ҫ������Ʊ���ȿ�����</font><br><br>
				</td> 
			</tr> 
			<tr><td align=center bgcolor="#E4E8EF"><input type="submit" name="submit" value="����" ></td></tr>
			</form>
		<% end if %>	
		</table>	
	</td>
	
	<td width=5></td>		<!--��һ�к͵ڶ���֮��ķָ���-->
	
    <td width="40%" align=right valign=top> 					<!-- �ڶ���--��ʵʱ����  -->
	
		<table cellspacing=1 cellpadding=0 border=0 width="100%" bgcolor="0066CC">
			<tr> 
				<td height=25 background='images/lan.gif' align=center><font color="#FFFFFF">����ʵʱ����</font></td>
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
				<td align=center background="images/footer.gif" height=25 bgcolor="#E4E8EF"><a href="news.asp" class="cblue" title="�鿴��������">��������</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="PStockApply.asp" class="cblue">��������</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="javascript:window.close()" class="cblue">�뿪��Ʊ��������</a></td>
			</tr>
		</table>

	</td>
		  
	<td width=5></td>		<!--�ڶ��к͵�����֮��ķָ���-->
		  
	<td width="30%" align="right" valign="top"> 	<!--������-->
	
		<table cellspacing=1 cellpadding=0 border=0 width="100%" bgcolor="0066CC">
			<tr align=center> 
				<td height=25 background='images/lan.gif'> <font color="#FFFFFF">�������ʲ����а�</font> </td>
			</tr>
			<tr bgcolor="#E4E8EF"> 
				<td> 
					<table height=22 cellspacing=0 cellpadding=1 width="100%" border=0>	
<%
					set rs=conn.execute("select top 12 �ʺ�,���ʽ�,id from [�ͻ�] order by [���ʽ�] desc")
					if RS.Eof then 
						response.Write "<td colspan=4> ��δ���������Ʊ </td>"
					else 
						i=0
						response.Write "<tr><td colspan=""4"" height=""10""></td></tr>"
						do while Not RS.Eof
							i=i+1
							response.Write "<tr><td align=""center"" width=""40""><font color=red>"&i&"</font></td><td width=""5""></td>"	
							response.Write "<td><a href='dispu.asp?uid="&rs(2)&"&username="&rs(0)&"'>"&rs(0)&"</a></td>"&_
										"<td align=right>"
							if rs(1)<=0 then 
								response.Write "0 Ԫ"
							else
								response.Write formatnumber(rs(1),0,-1)&" Ԫ&nbsp;"
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
				<td colspan=4 align=center background="images/title.gif" height=21 bgcolor="#E4E8EF" valign="middle"><a href="TopUser.asp?action=Reflash" class="cblue">ˢ������</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="TopUser.asp" class="cblue">��ϸ����</a></td>
			</tr>			
		</table>
	</td>
	</tr>
</table>

<% if master then %>
<br>
<table border="0" width="98%" cellspacing="1" cellpadding="3" align="center" bgcolor="#0066CC">
	<tr> 
		<td height=23 background='images/header.gif'>&nbsp;<font color="#FF0000">����ѡ��</font></td>
	</tr>
	<tr>
		<td height="26" valign="middle" bgcolor="#E4E8EF"><font color="#0066FF">�� <a href="Admin_Price.asp?action=UpPrice" title="�������й�Ʊ���ּ�">ȫ��</a> �� <a href="Admin_Price.asp?action=DownPrice" title="�������й�Ʊ���ּ�">ȫ��</a> �� <a href="Admin_Setting.asp">��������</a> �� <a href="Admin_Gupiao.asp">��Ʊ����</a> �� <a href="Admin_Gupiao.asp?action=newgupiao">�¹�����</a> �� <a href="Admin_User.asp">�ͻ�����</a> �� <a href="Admin_Data.asp">���ݿ����</a> ��</font></td>      
	</tr>
</table> 
<%end if%>  
<br>
<table cellspacing=1 cellpadding=0 width="98%" border=0 bgcolor="#0066CC" align=center>
	<tr><td height="23" background='images/lan.gif'><font color="#FFFFFF">&nbsp;���߹���</font></td></tr>
	<%call onlineuser(1,1)%>
</table> 
<br>
<table cellspacing=1 cellpadding=2 width="98%" border=0 bgcolor="#0066CC" align=center>
	<tr>  
		<td  height=32 background="images/footer.gif" valign=top>
			<table border=0 width="100%" cellspacing=0 cellpadding=5 bordercolorlight="#ffffff" bordercolordark="#ffffff">
				<tr valign="bottom"><td id="clock" width="185" align="left">��ӭ��</td><td width="*" id="CopyRight" align="center">�����޸ģ�һ����</td><td width="170" align="right">ҳ��ִ��ʱ�䣺<%=FormatNumber((timer()-startime)*1000,3)%>����</td></tr>
			</table>	
		</td>
	</tr>
</table> 
<% 
CloseDatabase		'�رչ�Ʊ���ݿ�

%>
</body>
</html>
<!--#include file="inc/js.asp"-->