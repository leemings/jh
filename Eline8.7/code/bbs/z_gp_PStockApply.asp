<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_gp_conn.asp"-->
<!--#include file="z_gp_Const.asp"-->
<%
stats="��������"
call nav()
call head_var(0,0,GuPiao_Setting(5),"z_gp_gupiao.asp")
dim PStock_Setting
PStock_Setting=gp_conn.execute("select top 1 PStock_Setting from GupiaoConfig order by id")(0)
PStock_Setting=split(PStock_Setting,"|")
PStock_Setting(0)=CLNG(PStock_Setting(0))
PStock_Setting(1)=CLNG(PStock_Setting(1))
PStock_Setting(2)=CLNG(PStock_Setting(2))
PStock_Setting(3)=CLNG(PStock_Setting(3))
PStock_Setting(4)=CLNG(PStock_Setting(4))
PStock_Setting(5)=CLNG(PStock_Setting(5))%>
<table class=tableborder1 cellspacing=1 cellpadding=3 align=center border=0 width="<%=Forum_body(12)%>">
<%select case request("action")
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
end select%>
</table>
<%call activeonline()
call footer()
'=====================================
sub main()%>
	<tr>
		<th valign=center align=middle height=25 colspan="8"><b>���˹�Ʊ��������һ����</b> | <a href="?action=apply" class="cblue">�����������</a></th>
	</tr>
	<tr align="center" height=20 valign="middle"> 
		<td class=tablebody2 align=center>���й�˾</td>
		<td class=tablebody2 align=center>��Ʊ����</td>
		<td class=tablebody2 align=center>���м۸�</td>
		<td class=tablebody2 align=center>������</td>
		<td class=tablebody2 align=center>��������</td>
		<td class=tablebody2 align=center>��˾���</td>
		<td class=tablebody2 align=center>����</td>
		<td class=tablebody2 align=center>����</td>
	</tr>
	<%dim currentpage,page_count,Pcount
	dim totalrec,endpage,explainsplit
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
	rs.open sql,gp_conn,1,1
	if rs.eof and rs.bof then
		response.write "<tr><td class=tablebody1 colspan=8>��ʱû��������������й�Ʊ</td></tr>"
	else
		rs.PageSize = Gupiao_Setting(2)
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		do while (not rs.eof) and (not page_count = 20)
			explainsplit=split(rs("explain"),"|")%>
		  <tr align="center"> 
				<td class=tablebody1><a href="?action=edit&sid=<%=rs("SID")%>"><%=rs("Stockname")%></a></td>
				<td class=tablebody1><%=rs("Stocknum")%></td>
				<td class=tablebody1 align=right><%=formatnumber(rs("Price"),2,true)%>&nbsp;</td>
				<td class=tablebody1><a href="z_gp_Dispu.asp?uid=<%=rs("uid")%>&username=<%=rs("Username")%>" class="cblue"><%=rs("Username")%></a></td>
				<td class=tablebody1><%=formatdatetime(rs("ApplyDate"),1)%></td>
				<td class=tablebody1><%=ExplainSplit(0)%></td>
				<td class=tablebody1><%if rs("States")=0 then%>δ���<%elseif rs("States")=1 then%>����׼<%else%>����׼<%end if%></td>
				<td class=tablebody1><%
					if master then
						if rs("States")=0 then
							%><a href="?action=ChangStat&stat=pass&sid=<%=rs("sid")%>">��׼</a> | <a href="?action=ChangStat&stat=nopass&sid=<%=rs("sid")%>">����׼</a><%elseif rs("States")=1 then%><a href="?action=ChangStat&stat=nopass&sid=<%=rs("sid")%>" class=cred>ժ��</a><%else%><a href="?action=ChangStat&stat=pass&sid=<%=rs("sid")%>">��׼</a><%end if%> | <a href="?action=edit&sid=<%=rs("sid")%>">�༭</a> | <a href="?action=del&sid=<%=rs("sid")%>" onclick="javascript:{if(confirm('��ȷ��ɾ�� <%=rs("Stockname")%> �����Ʊ��?')){return true;}return false;}" class=cred>ɾ��</a><%
						elseif rs("Uid")=MyUserID then
							%><a href="?action=edit&sid=<%=rs("sid")%>">�༭</a> <%if rs("States")=0 then%>| <a href="?action=del&reaction=my&sid=<%=rs("sid")%>" onclick="javascript:{if(confirm('��ȷ��ȡ�� <%=rs("Stockname")%> �����Ʊ������������?')){return true;}return false;}" class=cred>ע��</a><%
						end if
					else
						%>-<%
					end if
				%></td>   
	  	</tr>
			<%page_count = page_count + 1		
			rs.movenext
		loop
		Pcount=rs.PageCount%>
		<tr>
			<td class=tablebody2 colspan=2>���� <font color=blue><%=totalrec%></font> �����룬ÿҳ <font color=blue><%=rs.PageSize%></font> ������ <font color=red><%=currentpage%></font> ҳ / �� <font color=blue><%=Pcount%></font> ҳ</td>
			<td class=tablebody2 align=right colspan=6>��ҳ��<%call disppagenum(currentpage,pcount,"""?page=","""")%>&nbsp;&nbsp;&nbsp;&nbsp;</td>
		</tr>
	<%end if
	rs.close
	set rs=nothing%>			
	<tr>
		<th height=25 valign=middle align="center" colspan="8"><a href="?action=apply" class="cblue">[��������]</a>����<A href="z_gp_Gupiao.asp" class="cblue">[���ش���]</A></th>
	</tr>
<%end sub
'-----------------------
sub ChangStat()
	if not master then
		errmess="<li>��û��ִ�б�������Ȩ�ޣ�"
		call endinfo(1)
	else
		'on error resume next
		if request("stat")="pass" then
			set rs=gp_conn.execute("select * from PersonalStock where sid="&request("sid"))
			if not(rs.eof and rs.bof) then
				dim temp,GpID,PsMoney
				if CLNG(rs("States"))=0 then
					if CLNG(rs("Stocknum")*0.5)>10000 then
						temp="10000"
					elseif CLNG(rs("Stocknum")*0.5)>0 then
						temp=CLNG(rs("Stocknum")*0.5)
					else
						temp=0	
					end if				
					gp_conn.execute("insert into [GuPiao](QiYe,ZongGuFen,ShengYuGuFen,JiaoYiLiang,KaiPanJiaGe,DangQianJiaGe,ZhuangTai,RiQi,Explain,LtdImg,TongJi,JingYingZhe,Uid)"&_
					     " values('"&rs("Stockname")&"',"&rs("Stocknum")&","&clng(rs("Stocknum")*0.5)&","&temp&","&rs("Price")&","&rs("Price")&",'��',date(),'"&replace(rs("Explain"),"'","")&"|"&temp&"','"&rs("LtdImg")&"','0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0','"&rs("Username")&"',"&rs("Uid")&")")
					GpID=gp_conn.execute("select top 1 sid from [GuPiao] order by sid desc")(0)
					gp_conn.execute("insert into [DaHu] (Uid,ZhangHao,sid,MaiRuJiaGe,PingJunJiaGe,ChiGuShu,QiYe) values ("&rs("uid")&",'"&rs("Username")&"','"&GpID&"',"&rs("Price")&","&rs("Price")&","&clng(rs("Stocknum")*0.5)&",'"&rs("Stockname")&"')")
					PsMoney=rs("Stocknum")*0.5*rs("Price")
					gp_conn.execute("update [KeHu] set ZiJin=ZiJin-"&PsMoney*1.02&",ZongZiJin=ZongZiJin+"&PsMoney*0.98&",ChiGuZhongLei=ChiGuZhongLei+1,JinRiMaiRu=JinRiMaiRu+1,ZongMaiRu=ZongMaiRu+1 where ID="&rs("uid"))
					call ResetCid()
				elseif CLNG(rs("States"))=2 then
					gp_conn.execute("update [GuPiao] set [ZhuangTai]='��' where [QiYe]='"&rs(0)&"'")
				end if
				gp_conn.execute("update [PersonalStock] set States=1 where sid="&request("sid"))
			end if		
			rs.close
		elseif request("stat")="nopass" then
			set rs=gp_conn.execute("select Stockname,Price,Stocknum,States from [PersonalStock] where sid="&request("sid"))
			if not(rs.eof and rs.bof) then
				if CLNG(rs(3))=1 then
					gp_conn.execute("update [GuPiao] set [ZhuangTai]='��' where [QiYe]='"&rs(0)&"'")
				end if
				gp_conn.execute("update PersonalStock set States=2 where sid="&request("sid"))
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
	set rs=gp_conn.execute("select * from PersonalStock where sid="&Sid)
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
	end if%>
	<tr>
		<th colspan="2" valign=center align=middle height=25>�༭ �������й�˾ <b><%=rs("Stockname")%></b> ������</th>
	</tr>
	<form method=post  action="?action=SaveEdit">
	<input type=hidden name=sid value="<%=rs("sid")%>">
	<tr>
		<td class=tablebody1 width="50%"><b>���й�˾���ƣ�</b><br>����Ʊ�������3Byte���14Byte</td>
		<td class=tablebody1 width="50%">&nbsp;<input type=text name=gpname <%if rs("States")<>0 then%> readonly <%end if%> value="<%=rs("Stockname")%>"></td>
	</tr>
	<tr>
		<td class=tablebody1><b>���й�Ʊ���ۣ�</b></td>
		<td class=tablebody1>&nbsp;<%=rs("Price")%></td>
	</tr>	
	<tr>
		<td class=tablebody1><b>���й�Ʊ������</b></td>
		<td class=tablebody1>&nbsp;<%=rs("Stocknum")%> ��</td>  
	</tr>			
	<tr>
		<td class=tablebody1 valign="top"><b>���й�˾��飺</b><br>��Զ����й�˾����Ҫ�Ľ��ܣ�����������ĳɹ��ʻ���Ӱ��<br>��˾��鲻�ܳ���100Byte</td>  
		<td class=tablebody1>&nbsp;<textarea cols="39" name="Explain" rows="5" wrap="PHYSICAL"><%=rs("Explain")%></textarea></td>  
	</tr>
	<tr>
		<td class=tablebody2 colspan="2" align="center"><input type=submit name=submit value=�ύ>��<input type=reset name=submit value=����></td> 
	</tr>
	</form>
	<tr>
		<th height=25 valign=middle align="center" colspan="2"><A href="z_gp_Gupiao.asp">[���ع���]</A>����<A href="<%=Request.ServerVariables("HTTP_REFERER")%>">[������һҳ]</A></th>
	</tr>
	<%rs.close 
end sub 
'-------------�����޸�---------------
sub SaveEdit() 
	dim GpName,Explain
	GpName=trim(replace(Request.form("gpname"),"'","")) 
	Explain=trim(replace(Request.form("Explain"),"'","")) 

	if GpName="" then
		errmess="<li>���������й�˾������"
		founderr=true
	elseif len(GpName)>PStock_Setting(3) or len(GpName)<PStock_Setting(2) then
		errmess="<li>���й�˾���Ƶĳ��Ȳ��ܳ��� "&PStock_Setting(3)&" Byte����С�� "&PStock_Setting(2)&" Byte"
		founderr=true
	end if 
	if len(Explain)>PStock_Setting(4) then
		errmess=errmess+"<li>��˾��鲻�ܳ���"&PStock_Setting(4)&"Byte"
		founderr=true
	end if		

	if founderr then
		call endinfo(2)
		exit sub	
	end if	
	if Explain="" then Explain="����һ������ʲô��û�����£�"
	gp_conn.execute("update PersonalStock set Stockname='"&GpName&"',Explain='"&Request.form("Explain")&"' where sid="&request.form("sid"))
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
		set rs=gp_conn.execute("select Stockname,Username,Uid from PersonalStock where sid="&DelID)
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
		gp_conn.execute("delete from PersonalStock where sid="&DelID)
		set rs=gp_conn.execute("select * from [DaHu] where sid="&DelID)
		
		if request("reaction")<>"my" then
			sucmess="<font color=#66CCFF>�����й�˾ "&GPname&"</font><font color=#CCCCCC> δ�ܻ�ù�Ʊ֤�����׼����������</font>"
			gp_conn.execute "insert into RndEvent(content,AddTime) values('"&sucmess&"','"&now()&"' )"  
			sucmess="<li>"&GPname&" ���й�˾ɾ���ɹ�"
		else
			sucmess="<li>���Ѿ��ɹ��Ѹ������е�����ע����"	
		end if
		call endinfo(2)
	end if	
end sub
'-----------------�¹�����---------------
sub apply()
	if PStock_Setting(0)=0 then
		errmess="<li>����������ͣ����"
		call endinfo(1)
		exit sub		
	end if
	if MyCash<csng(PStock_Setting(1)) then
		errmess="<li>�����ڵĹ�Ʊ�ʻ��ʽ��� "&PStock_Setting(1)&" Ԫ����������������й�˾"
		call endinfo(1)
		exit sub
	end if
	
	dim MinPrice,MaxPrice,YourCash
	set rs=gp_conn.execute("select top 1 KaiPanJiaGe from [GuPiao] order by KaiPanJiaGe")
	if rs.eof then 
		MinPrice=2.00
	elseif rs(0)*0.5<2 then
		MinPrice=2.00
	else
		MinPrice=formatnumber(rs(0)*0.5,2,true)
	end if	
	set rs=gp_conn.execute("select top 1 KaiPanJiaGe from [GuPiao] order by KaiPanJiaGe desc")
	if rs.eof then 
		MaxPrice=3.00
	elseif rs(0)*0.5<3 then
		MaxPrice=3.00
	else
		MaxPrice=formatnumber(rs(0)*0.5,2,true)	
	end if	
	YourCash=gp_conn.execute("select ZiJin from [KeHu] where id="&MyUserID&" and suoding<2")(0)%>
	<tr>
		<th colspan="2" valign=center align=middle height=25><b>�������������</b></th>
	</tr>
	<form method=post action="?action=SaveApply" name="form1">
	<tr>
		<td class=tablebody1 width="50%"><b>�����ˣ�</b></td>
		<td class=tablebody1 width="50%">&nbsp;<%=membername%></td>
	</tr>
	<tr>
		<td class=tablebody1 width="50%"><b>���й�˾���ƣ�</b><br>����Ʊ�����<%=PStock_Setting(3)%>Byte�����<%=PStock_Setting(2)%>Byte</td>
		<td class=tablebody1 width="50%">&nbsp;<input type=text name=gpname value=""> <font color=red>*</font></td>
	</tr>
	<tr>
		<td class=tablebody1><b>���й�Ʊ���ۣ�</b><br>���۷�Χ��<%=MinPrice%>��<%=MaxPrice%></td>
		<td class=tablebody1>&nbsp;<input type=text name=price value=""> <font color=red>*</font> <input type=button value='����������й�Ʊ����' name=Button onclick=checknum(<%=MinPrice%>,<%=MaxPrice%>)></td>
	</tr>	
	<tr>
		<td class=tablebody1><b>���й�Ʊ����</b><br>ÿ�ɵ��۳��������㷢�еĹ�Ʊ�����ܴ��������ڵ��˻��ʽ��ܶ�</td>
		<td class=tablebody1>&nbsp;<input type=text name=gpnum value=""> <font color=red>*</font> <span id="maxnum"></span></td>
	</tr>			
	<tr>
		<td class=tablebody1 valign="top"><b>���й�˾��飺</b><br>������й�˾����Ҫ�Ľ��ܣ�����������ĳɹ��ʻ���Ӱ��<br>��˾��鲻�ܳ���<%=PStock_Setting(4)%>Byte</td>   
		<td class=tablebody1>&nbsp;<textarea cols="39" name="Explain" rows="5" wrap="PHYSICAL">����һ������ʲô��û�����£�</textarea></td> 
	</tr>
	<tr>
		<td class=tablebody2 colspan="2" align="center"><input type=submit name=submit value=����>��<input type=reset name=submit value=����></td>
	</tr>
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
	<tr>
		<th height=25 valign=middle align="center" colspan="2"><A href="z_gp_Gupiao.asp">[���ع���]</A>����<A href="<%=Request.ServerVariables("HTTP_REFERER")%>">[������һҳ]</A></th>
	</tr>
<%end sub
'-------------�����¹�Ʊ---------------
sub SaveApply()
	if PStock_Setting(0)=0 then
		errmess="<li>����������ͣ����"
		call endinfo(1)
		exit sub		
	end if
	if MyCash<csng(PStock_Setting(1)) then
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
	elseif len(GpName)>PStock_Setting(3) or len(GpName)<PStock_Setting(2) then
		errmess="<li>���й�˾���Ƶĳ��Ȳ��ܳ��� "&PStock_Setting(3)&" Byte����С�� "&PStock_Setting(2)&" Byte"
		founderr=true
	end if 
	if len(Explain)>PStock_Setting(4) then
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
		MaxGpNum=clng(Request.form("YourCash")/Price)
		if gpnum="" or not(isnumeric(gpnum)) then
			errmess=errmess+"<br><li>���������й�Ʊ��������������Ĳ�������"
			founderr=true
		elseif clng(gpnum)<clng(PStock_Setting(5)) then
			errmess=errmess+"<br><li>��Ʊ����������� "&PStock_Setting(5)&" ��"
			founderr=true		
		elseif clng(gpnum)>clng(MaxGpNum) then
			errmess=errmess+"<br><li>�����������еĹ�Ʊ������ "&MaxGpNum&" ��"
			founderr=true
		end if
	end if	
	
	if founderr then
		call endinfo(2)
		exit sub
	end if
			
	on error resume next 
	set rs=gp_conn.execute("select sid from [GuPiao] where QiYe='"&GpName&"'")
	if not rs.eof then
		errmess="<li>����������й�˾�Ѿ����ڣ���������д���й�˾����"
		endinfo(2)
		exit sub
	end if
	set rs=gp_conn.execute("select Uid from PersonalStock where Stockname='"&GpName&"'")
	if not rs.eof then
		errmess="<li>����������й�˾�Ѿ������ˣ���������д���й�˾����"
		endinfo(2)
		exit sub
	end if
	set rs=gp_conn.execute("select Uid from PersonalStock where Uid="&MyUserID&"")	
	if not rs.eof then
		errmess="<li>���Ѿ����Լ������й�˾��"
		endinfo(2)
		exit sub		
	end if
			
	gp_conn.execute("insert into PersonalStock(Username,Uid,Stockname,Price,Stocknum,Explain,ApplyDate,States) values('"&membername&"',"&MyUserID&",'"&GpName&"',"&Price&","&gpnum&",'"&Explain&"',now(),0)")
	
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
    set rs=gp_conn.execute("select sid from [GuPiao] order by sid")
	Cid=1
	do while not rs.eof
		gp_conn.execute("update [GuPiao] set Cid="&Cid&" where sid="&rs(0))
		Cid=Cid+1	
		rs.movenext
	loop
end sub
%>