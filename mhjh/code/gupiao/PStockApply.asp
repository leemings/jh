<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html><head><title><%=Gupiao_Setting(5)%>-��������</title>
<!--#include file="css.asp"-->
</head><body bgcolor="#ffffff" text="#000000" style="FONT-SIZE: 9pt" topmargin=5 leftmargin=0 oncontextmenu=self.event.returnValue=false>
<% 
	dim PStock_Setting
	PStock_Setting=conn.execute("select top 1 PStock_Setting from GupiaoConfig order by id")(0)
	PStock_Setting=split(PStock_Setting,"|")
	
	select case request("action")
		case "edit"
			call Edit()
		case "SaveEdit"		
			call SaveEdit()
		case "del"
			call del()
		case "apply"
			call apply()
		case "SaveApply"		
			call SaveApply()
		case "ChangStat"
			call ChangStat()					
		case else
			call main()
	end select
%>
</body></html>
<%
CloseDatabase		'�ر����ݿ�
'=====================================
sub main()
%>

<table width="97%" border="0" cellspacing="1" cellpadding="3" align="center" bgcolor="#0066CC">
	<TR>
		<TD BACKGROUND="Images/topbg.gif" height=9 colspan=8></TD>
	</TR>
	<tr>
		<td valign=center align=middle height=23 background="Images/Header.gif" colspan="8"><b>���˹�Ʊ��������һ����</b> | <a href="?action=apply" class="cblue">�����������</a></td>
	</tr>
	<tr align="center" height=20 valign="middle"> 
		<td background="Images/lan.gif"><font color=#ffffff>���й�˾</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>��Ʊ����</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>���м۸�</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>������</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>��������</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>��˾���</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>����</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>����</font></td>
	</tr>
<%        
	dim currentpage,page_count,Pcount
	dim totalrec,endpage
	currentPage=request("page")
	if currentpage="" or not isnumeric(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
		if err then
			currentpage=1
			err.clear
		end if
	end if
	Set rs= Server.CreateObject("ADODB.Recordset")
	sql= "select * from PersonalStock"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "<tr><td colspan=8 bgcolor=#E4E8EF>��ʱû��������������й�Ʊ</td></tr>"
	else
		rs.PageSize = Gupiao_Setting(2)
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		do while (not rs.eof) and (not page_count = 20)	
%>
	  <tr align="center"> 
		<td bgcolor="#FFFFFF"><a href="?action=edit&sid=<%=rs("SID")%>"><%=rs("Stockname")%></a></td>
		<td bgcolor="#FFFFFF"><%=rs("Stocknum")%></td>
		<td bgcolor="#FFFFFF"><%=formatcurrency(rs("Price"),3,-1)%></td>
		<td bgcolor="#FFFFFF"><a href="dispu.asp?uid=<%=rs("uid")%>&username=<%=rs("Username")%>" class="cblue"><%=rs("Username")%></a></td>
		<td bgcolor="#FFFFFF"><%=formatdatetime(rs("ApplyDate"),1)%></td>
		<td bgcolor="#FFFFFF"><%=rs("Explain")%></td>
		<td bgcolor="#FFFFFF"><%if rs("States")=0 then%>δ���<%elseif rs("States")=1 then%>����׼<%else%>����׼<%end if%></td>
		<td bgcolor="#FFFFFF">
		<%if master then%>
			<%if rs("States")=0 then%><a href="?action=ChangStat&stat=pass&sid=<%=rs("sid")%>">��׼</a> | <a href="?action=ChangStat&stat=nopass&sid=<%=rs("sid")%>">����׼</a><%elseif rs("States")=1 then%><a href="?action=ChangStat&stat=nopass&sid=<%=rs("sid")%>" class=cred>ժ��</a><%else%><a href="?action=ChangStat&stat=pass&sid=<%=rs("sid")%>">��׼</a><%end if%> | <a href="?action=edit&sid=<%=rs("sid")%>">�༭</a> | <a href="?action=del&sid=<%=rs("sid")%>" onclick="javascript:{if(confirm('��ȷ��ɾ�� <%=rs("Stockname")%> �����Ʊ��?')){return true;}return false;}" class=cred>ɾ��</a> 
		<%elseif rs("Uid")=MyUserID then%>
			<a href="?action=edit&sid=<%=rs("sid")%>">�༭</a> <%if rs("States")=0 then%>| <a href="?action=del&reaction=my&sid=<%=rs("sid")%>" onclick="javascript:{if(confirm('��ȷ��ȡ�� <%=rs("Stockname")%> �����Ʊ������������?')){return true;}return false;}" class=cred>ע��</a><%end if%>
		<%else%>
			-	
		<%end if%>
		</td>   
	  </tr>
<%         
		page_count = page_count + 1		
		rs.movenext
	loop
	Pcount=rs.PageCount         
%>
	<tr><td colspan=8 background="images/title.gif">
		<table border="0" cellpadding="0" cellspacing="0" width="100%"><tr>
			<td align=left>����<font color=blue><%=totalrec%></font>�����룬ÿҳ<font color=blue><%=Gupiao_Setting(2)%></font>������<font color=red><%=currentpage%></font>ҳ/��<font color=blue><%=Pcount%></font>ҳ</td>
			<td align=right>��ҳ��
<%
			if currentpage > 4 then
				response.write "<a href=""?page=1"">[1]</a> ..."
			end if
			if Pcount>currentpage+3 then
			endpage=currentpage+3
			else
			endpage=Pcount
			end if
			for i=currentpage-3 to endpage
			if not i<1 then
				if i = clng(currentpage) then
				response.write " <font color=red>["&i&"]</font>"
				else
				response.write " <a href=""?page="&i&""">["&i&"]</a>"
				end if
			end if
			next
			if currentpage+3 < Pcount then 
				response.write "... <a href=""?page="&Pcount&""">["&Pcount&"]</a>"
			end if
%>
				</td></tr>
			 </table>
		</td></tr>
<%				
		end if
		rs.close
		set rs=nothing	
%>			
	<tr><td height=32 background="images/footer.gif" valign=middle align="center" colspan="8"><a href="?action=apply" class="cblue">[��������]</a>����<A href="stock.asp" class="cblue">[���ش���]</A></td></tr>
</table>
<% 
end sub
'-----------------------
sub ChangStat()
	if not master then
		errmess="<li>��û��ִ�б�������Ȩ�ޣ�"
		call endinfo(1)
	else
		'on error resume next
		if request("stat")="pass" then
			set rs=conn.execute("select * from PersonalStock where sid="&request("sid"))
			if not(rs.eof and rs.bof) then
				dim temp,GpID,PsMoney
				if cint(rs("States"))=0 then
					if clng(rs("Stocknum")*0.5)>10000 then
						temp="10000"
					elseif clng(rs("Stocknum")*0.5)>0 then
						temp=clng(rs("Stocknum")*0.5)
					else
						temp=0	
					end if				
					conn.execute("insert into [��Ʊ](��ҵ,�ܹɷ�,ʣ��ɷ�,������,IniTradeNum,���̼۸�,��ǰ�۸�,״̬,����,Explain,LtdImg,TongJi,��Ӫ��,Uid)"&_
					     " values('"&rs("Stockname")&"',"&rs("Stocknum")&","&clng(rs("Stocknum")*0.5)&","&temp&","&temp&","&rs("Price")&","&rs("Price")&",'��',date(),'"&replace(rs("Explain"),"'","")&"','"&rs("LtdImg")&"','0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0','"&rs("Username")&"',"&rs("Uid")&")")
					GpID=conn.execute("select top 1 sid from [��Ʊ] order by sid desc")(0)
					conn.execute("insert into [��] (Uid,�ʺ�,sid,����۸�,ƽ���۸�,�ֹ���,��ҵ) values ("&rs("uid")&",'"&rs("Username")&"','"&GpID&"',"&rs("Price")&","&rs("Price")&","&clng(rs("Stocknum")*0.5)&",'"&rs("Stockname")&"')")
					PsMoney=rs("Stocknum")*0.5*rs("Price")
					conn.execute("update [�ͻ�] set �ʽ�=�ʽ�-"&PsMoney*1.02&",���ʽ�=���ʽ�+"&PsMoney*0.98&",�ֹ�����=�ֹ�����+1,��������=��������+1,������=������+1 where ID="&rs("uid"))
					call ResetCid()
				elseif cint(rs("States"))=2 then
					conn.execute("update [��Ʊ] set [״̬]='��' where [��ҵ]='"&rs(0)&"'")
				end if
				conn.execute("update [PersonalStock] set States=1 where sid="&request("sid"))
			end if		
			rs.close
		elseif request("stat")="nopass" then
			set rs=conn.execute("select Stockname,Price,Stocknum,States from [PersonalStock] where sid="&request("sid"))
			if not(rs.eof and rs.bof) then
				if cint(rs(3))=1 then
					conn.execute("update [��Ʊ] set [״̬]='��' where [��ҵ]='"&rs(0)&"'")
				end if
				conn.execute("update PersonalStock set States=2 where sid="&request("sid"))
			end if
			rs.close		
		end if
		if err.number<>"0" then
			errmess="<li>���ִ���"&Err.Description
			call endinfo(1)
			exit sub		
		end if
		response.redirect "?"
	end if	
end sub
'-------------�༭��Ʊ---------------
sub Edit()
	dim Sid
	Sid=request("sid")
	if Sid="" then
		errmess="<li>����������ָ����ع�Ʊ"
		call endinfo(2)
		exit sub	
	end if
	set rs=conn.execute("select * from PersonalStock where sid="&Sid)
	if rs.eof and rs.bof then
		rs.close
		errmess="<li>û���ҵ�ָ�������й�˾"
		call endinfo(2)
		exit sub
	elseif rs("Uid")<>MyUserID then
		rs.close
		errmess="<li>�����Ǹ����й�˾�Ĵ�ɶ���û��ִ�б�������Ȩ��"
		call endinfo(2)
		exit sub		
	end if
%>
<table  width="550" height=300 border=0 cellspacing=1 cellpadding=3 align=center bgcolor="#0066CC">
	<TR>
	<TD BACKGROUND="images/topbg.gif" height=9 colspan="2"></TD>
	</TR>
	<tr>
		<td colspan="2" valign=center align=middle height=23 class=big background="images/header.gif">�༭ �������й�˾ <b><%=rs("Stockname")%></b> ������</td>
	</tr>
	<form method=post  action="?action=SaveEdit">
	<input type=hidden name=sid value="<%=rs("sid")%>">
	<tr>
		<td BGCOLOR="#E4E4EF" width="50%"><b>���й�˾���ƣ�</b><br>����Ʊ�������3Byte���14Byte</td>
		<td bgcolor="#FFFFFF" width="50%">&nbsp;<input type=text name=gpname <%if rs("States")<>0 then%> readonly <%end if%> value="<%=rs("Stockname")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF"><b>���й�Ʊ���ۣ�</b></td>
		<td bgcolor="#FFFFFF">&nbsp;<%=rs("Price")%> ��</td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF"><b>���й�Ʊ������</b></td>
		<td bgcolor="#FFFFFF">&nbsp;<%=rs("Stocknum")%> ��</td>  
	</tr>			
	<tr>
		<td BGCOLOR="#E4E4EF" valign="top"><b>���й�˾��飺</b><br>��Զ����й�˾����Ҫ�Ľ��ܣ�����������ĳɹ��ʻ���Ӱ��<br>��˾��鲻�ܳ���100Byte</td>  
		<td bgcolor="#FFFFFF">&nbsp;<textarea cols="39" name="Explain" rows="5" wrap="PHYSICAL"><%=rs("Explain")%></textarea></td>  
	</tr>
	<tr><td colspan="2" bgcolor="#FFFFFF" align="center"><input type=submit name=submit value=�ύ>��<input type=reset name=submit value=����></td> 
	</form>
	<tr><td height=32 background="images/footer.gif" valign=middle align="center" colspan="8"><A href="stock.asp">[���ع���]</A>����<A href="<%=Request.ServerVariables("HTTP_REFERER")%>">[������һҳ]</A></td></tr>
</table>
<%	
	rs.close 
end sub 
'-------------�����޸�---------------
sub SaveEdit() 
	dim GpName,Explain
	GpName=trim(replace(Request.form("gpname"),"'","")) 
	Explain=trim(replace(Request.form("Explain"),"'","")) 

	if GpName="" then
		errmess="<li>���������й�˾������"
		founderr=true
	elseif len(GpName)>cint(PStock_Setting(3)) or len(GpName)<cint(PStock_Setting(2)) then
		errmess="<li>���й�˾���Ƶĳ��Ȳ��ܳ��� "&PStock_Setting(3)&" Byte����С�� "&PStock_Setting(2)&" Byte"
		founderr=true
	end if 
	if len(Explain)>cint(PStock_Setting(4)) then
		errmess=errmess+"<li>��˾��鲻�ܳ���"&PStock_Setting(4)&"Byte"
		founderr=true
	end if		

	if founderr then
		call endinfo(2)
		exit sub	
	end if	
	if Explain="" then Explain="����һ������ʲô��û�����£�"
	conn.execute("update PersonalStock set Stockname='"&GpName&"',Explain='"&Request.form("Explain")&"' where sid="&request.form("sid"))
	sucmess="<li>���й�˾��Ϣ�Ѿ��޸ĳɹ���"
	rUrl="?"
	call endinfo(2)
end sub
'---------------ɾ����Ʊ-------------
sub del()
	if (not master) and request("reaction")<>"my" then
		errmess="<li>��û��ִ�б�������Ȩ�ޣ�"
		call endinfo(1)
	else
		dim DelID,GPname
		DelID=request("sid")
		if DelID="" then
			errmess="<li>����������ָ��Ҫɾ���Ĺ�Ʊ"
			call endinfo(2)
			exit sub	
		end if
		set rs=conn.execute("select Stockname,Username,Uid from PersonalStock where sid="&DelID)
		if rs.eof and rs.bof then
			rs.close
			errmess="<li>û���ҵ���Ҫɾ���Ĺ�Ʊ"
			call endinfo(2)
			exit sub
		elseif rs(2)<>MyUserID then
			errmess="<li>�Ƿ���������û��ִ�иò�����Ȩ��"
			call endinfo(2)
			exit sub
		else			
			GPname=rs(0)
		end if
		rs.close
		conn.execute("delete from PersonalStock where sid="&DelID)
		set rs=conn.execute("select * from [��] where sid="&DelID)
		
		if request("reaction")<>"my" then
			sucmess="<font color=#66CCFF>�����й�˾ "&GPname&"</font><font color=#CCCCCC> δ�ܻ�ù�Ʊ֤�����׼����������</font>"
			conn.execute "insert into RndEvent(content,AddTime) values('"&sucmess&"','"&now()&"' )"  
			sucmess="<li>"&GPname&" ���й�˾ɾ���ɹ�"
		else
			sucmess="<li>���Ѿ��ɹ��Ѹ������е�����ע����"	
		end if
		call endinfo(2)
	end if	
end sub
'-----------------�¹�����---------------
sub apply()
	if cint(PStock_Setting(0))=0 then
		errmess="<li>����������ͣ����"
		call endinfo(1)
		exit sub		
	end if
	if MyCash<cdbl(PStock_Setting(1)) then
		errmess="<li>�����ڵĹ�Ʊ�ʻ��ʽ��� "&PStock_Setting(1)&" Ԫ����������������й�˾"
		call endinfo(1)
		exit sub
	end if
	
	dim MinPrice,MaxPrice,YourCash
	set rs=conn.execute("select top 1 ���̼۸� from [��Ʊ] order by ���̼۸�")
	if rs.eof then 
		MinPrice=1.000
	elseif rs(0)*0.5<1 then
		MinPrice=1.000
	else
		MinPrice=formatnumber(rs(0)*0.5,3,-1)
	end if	
	set rs=conn.execute("select top 1 ���̼۸� from [��Ʊ] order by ���̼۸� desc")
	if rs.eof then 
		MaxPrice=100.000
	elseif rs(0)*0.5<1 then
		MaxPrice=100.000
	else
		MaxPrice=formatnumber(rs(0)*0.5,3,-1)	
	end if	
	YourCash=conn.execute("select �ʽ� from [�ͻ�] where id="&MyUserID&"")(0)
%>
<table  width="80%" border=0 cellspacing=1 cellpadding=3 align=center bgcolor="#0066CC">
	<TR>
	<TD BACKGROUND="images/topbg.gif" height=9 colspan="2"></TD>
	</TR>
	<tr>
		<td colspan="2" valign=center align=middle height=23 class=big background="images/header.gif"><b>�������������</b></td>
	</tr>
	<form method=post action="?action=SaveApply" name="form1">
	<tr>
		<td BGCOLOR="#E4E4EF" width="50%"><b>�����ˣ�</b></td>
		<td bgcolor="#FFFFFF" width="50%">&nbsp;<%=membername%></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF" width="50%"><b>���й�˾���ƣ�</b><br>����Ʊ�����<%=PStock_Setting(3)%>Byte�����<%=PStock_Setting(2)%>Byte</td>
		<td bgcolor="#FFFFFF" width="50%">&nbsp;<input type=text name=gpname value=""> <font color=red>*</font></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF"><b>���й�Ʊ���ۣ�</b><br>���۷�Χ����<%=MinPrice%>����<%=MaxPrice%></td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=price value=""> <font color=red>*</font> <input type=button value='����������й�Ʊ����' name=Button onclick=checknum(<%=MinPrice%>,<%=MaxPrice%>)></td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF"><b>���й�Ʊ����</b><br>ÿ�ɵ��۳��������㷢�еĹ�Ʊ�����ܴ��������ڵ��˻��ʽ��ܶ�</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=gpnum value=""> <font color=red>*</font> <span id="maxnum"></span></td>
	</tr>			
	<tr>
		<td BGCOLOR="#E4E4EF" valign="top"><b>���й�˾��飺</b><br>��Զ����й�˾����Ҫ�Ľ��ܣ�����������ĳɹ��ʻ���Ӱ��<br>��˾��鲻�ܳ���<%=PStock_Setting(4)%>Byte</td>   
		<td bgcolor="#FFFFFF">&nbsp;<textarea cols="39" name="Explain" rows="5" wrap="PHYSICAL">����һ������ʲô��û�����£�</textarea></td> 
	</tr>
	<tr><td colspan="2" bgcolor="#FFFFFF" align="center"><input type=submit name=submit value=����>��<input type=reset name=submit value=����></td>
	<input type=hidden name=MinPrice value="<%=MinPrice%>">
	<input type=hidden name=MaxPrice value="<%=MaxPrice%>">
	<input type=hidden name=YourCash value="<%=YourCash%>">
	<script language="VBScript">
	function checknum(minp,maxp)
		dim price,YourCash
		price=document.form1.price.value
		YourCash=document.form1.YourCash.value
		if price="" then
			msgbox "�������Ʊ����",64,"��������"
			exit function
		elseif not isnumeric(price) then
			msgbox "��Ʊ���۱���������ֵ",64,"��������"
			exit function
		else
			price=price+0
		end if	
		if price<minp+0 then
			msgbox "��Ʊ���۲���С�� "&minp,64,"��������"
			exit function
		elseif price>maxp+0 then
			msgbox "��Ʊ���۲��ܴ��� "&maxp,64,"��������"
			exit function				
		else			
			maxnum.innerHTML="������������ <font color=blue>"&int(YourCash/price)&" </font> ��"
		end if
	end function
	</script>
	</form>
	<tr><td height=32 background="images/footer.gif" valign=middle align="center" colspan="8"><A href="stock.asp">[���ع���]</A>����<A href="<%=Request.ServerVariables("HTTP_REFERER")%>">[������һҳ]</A></td></tr>
</table>
<%
end sub
'-------------�����¹�Ʊ---------------
sub SaveApply()
	if cint(PStock_Setting(0))=0 then
		errmess="<li>����������ͣ����"
		call endinfo(1)
		exit sub		
	end if
	if MyCash<cdbl(PStock_Setting(1)) then
		errmess="<li>�����ڵĹ�Ʊ�ʻ��ʽ��� "&PStock_Setting(1)&" Ԫ����������������й�˾"
		call endinfo(1)
		exit sub
	end if		
	dim GpName,gpnum,Price,MinPrice,MaxPrice,MaxGpNum,Explain
	GpName=trim(replace(Request.form("gpname"),"'","")) 
	gpnum=Request.form("gpnum")
	Price=Request.form("price")
	MinPrice=Request.form("MinPrice")
	MaxPrice=Request.form("MaxPrice")
	Explain=trim(replace(Request.form("Explain"),"'","")) 
	
	if Explain="" then Explain="����һ������ʲô��û�����£�"
	if GpName="" then
		errmess="<li>���������й�˾������"
		founderr=true
	elseif len(GpName)>cint(PStock_Setting(3)) or len(GpName)<cint(PStock_Setting(2)) then
		errmess="<li>���й�˾���Ƶĳ��Ȳ��ܳ��� "&PStock_Setting(3)&" Byte����С�� "&PStock_Setting(2)&" Byte"
		founderr=true
	end if 
	if len(Explain)>cint(PStock_Setting(4)) then
		errmess=errmess+"<li>��˾��鲻�ܳ���"&PStock_Setting(4)&"Byte"
		founderr=true
	end if
	if Price="" or (not isnumeric(Price)) then
		errmess=errmess+"<br><li>���������й�Ʊ���ۻ���������Ĳ�������"
		founderr=true
	elseif Price+0>MaxPrice+0 then
		errmess=errmess+"<br><li>���й�Ʊ���۲��ܴ��� "&MaxPrice&" Ԫ"
		founderr=true
	elseif Price+0<MinPrice+0 then
		errmess=errmess+"<br><li>���й�Ʊ���۲���С�� "&MinPrice&" Ԫ"
		founderr=true		
	end if
	if not founderr then	
		MaxGpNum=cdbl(Request.form("YourCash")/Price)
		if gpnum="" or not(isnumeric(gpnum)) then
			errmess=errmess+"<br><li>���������й�Ʊ��������������Ĳ�������"
			founderr=true
		elseif cdbl(gpnum)<cdbl(PStock_Setting(5)) then
			errmess=errmess+"<br><li>��Ʊ����������� "&PStock_Setting(5)&" ��"
			founderr=true		
		elseif cdbl(gpnum)>cdbl(MaxGpNum) then
			errmess=errmess+"<br><li>�����������еĹ�Ʊ������ "&MaxGpNum&" ��"
			founderr=true
		end if
	end if	
	
	if founderr then
		call endinfo(2)
		exit sub
	end if
			
	on error resume next 
	set rs=conn.execute("select sid from ��Ʊ where ��ҵ='"&GpName&"'")
	if not rs.eof then
		errmess="<li>����������й�˾�Ѿ����ڣ���������д���й�˾����"
		endinfo(2)
		exit sub
	end if
	set rs=conn.execute("select Uid from PersonalStock where Stockname='"&GpName&"'")
	if not rs.eof then
		errmess="<li>����������й�˾�Ѿ������ˣ���������д���й�˾����"
		endinfo(2)
		exit sub
	end if
	set rs=conn.execute("select Uid from PersonalStock where Uid="&MyUserID&"")	
	if not rs.eof then
		errmess="<li>���Ѿ����Լ������й�˾��"
		endinfo(2)
		exit sub		
	end if
			
	conn.execute("insert into PersonalStock(Username,Uid,Stockname,Price,Stocknum,Explain,ApplyDate,States) values('"&membername&"',"&MyUserID&",'"&GpName&"',"&Price&","&gpnum&",'"&Explain&"',now(),0)")
	
	if err.number=0 then
		sucmess="<li>�����й�˾ <font color=blue>"&Request.form("gpname")&"</font> �ɹ�����"
		sucmess=sucmess+"<li>��ȴ�֤����������ͨ������֮��ͻ���ʽ����"
		rUrl="?"
	else
		errmess="<li>���ִ��󣬾������£�"&Err.Description
	end if	
	call endinfo(2)
end sub
'---------------������Ʊ��Cid--------------------
sub ResetCid()
'˵����Cid ����������¼��ģ������Ǵ�1��ʼ����Ȼ����1��2��3��4��5........
	Dim Cid 
    set rs=conn.execute("select sid from [��Ʊ] order by sid")
	Cid=1
	do while not rs.eof
		conn.execute("update [��Ʊ] set Cid="&Cid&" where sid="&rs(0))
		Cid=Cid+1	
		rs.movenext
	loop
end sub
%>