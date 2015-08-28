<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html><head><title><%=Gupiao_Setting(5)%>-个股上市</title>
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
CloseDatabase		'关闭数据库
'=====================================
sub main()
%>

<table width="97%" border="0" cellspacing="1" cellpadding="3" align="center" bgcolor="#0066CC">
	<TR>
		<TD BACKGROUND="Images/topbg.gif" height=9 colspan=8></TD>
	</TR>
	<tr>
		<td valign=center align=middle height=23 background="Images/Header.gif" colspan="8"><b>个人股票上市申请一览表</b> | <a href="?action=apply" class="cblue">申请个股上市</a></td>
	</tr>
	<tr align="center" height=20 valign="middle"> 
		<td background="Images/lan.gif"><font color=#ffffff>上市公司</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>股票数量</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>上市价格</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>申请人</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>申请日期</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>公司简介</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>审批</font></td>
		<td background="Images/lan.gif"><font color=#ffffff>操作</font></td>
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
		response.write "<tr><td colspan=8 bgcolor=#E4E8EF>暂时没有人申请个股上市股票</td></tr>"
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
		<td bgcolor="#FFFFFF"><%if rs("States")=0 then%>未审核<%elseif rs("States")=1 then%>已批准<%else%>不批准<%end if%></td>
		<td bgcolor="#FFFFFF">
		<%if master then%>
			<%if rs("States")=0 then%><a href="?action=ChangStat&stat=pass&sid=<%=rs("sid")%>">批准</a> | <a href="?action=ChangStat&stat=nopass&sid=<%=rs("sid")%>">不批准</a><%elseif rs("States")=1 then%><a href="?action=ChangStat&stat=nopass&sid=<%=rs("sid")%>" class=cred>摘牌</a><%else%><a href="?action=ChangStat&stat=pass&sid=<%=rs("sid")%>">批准</a><%end if%> | <a href="?action=edit&sid=<%=rs("sid")%>">编辑</a> | <a href="?action=del&sid=<%=rs("sid")%>" onclick="javascript:{if(confirm('您确定删除 <%=rs("Stockname")%> 这个股票吗?')){return true;}return false;}" class=cred>删除</a> 
		<%elseif rs("Uid")=MyUserID then%>
			<a href="?action=edit&sid=<%=rs("sid")%>">编辑</a> <%if rs("States")=0 then%>| <a href="?action=del&reaction=my&sid=<%=rs("sid")%>" onclick="javascript:{if(confirm('您确定取消 <%=rs("Stockname")%> 这个股票的上市申请吗?')){return true;}return false;}" class=cred>注销</a><%end if%>
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
			<td align=left>共有<font color=blue><%=totalrec%></font>个申请，每页<font color=blue><%=Gupiao_Setting(2)%></font>个，第<font color=red><%=currentpage%></font>页/共<font color=blue><%=Pcount%></font>页</td>
			<td align=right>分页：
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
	<tr><td height=32 background="images/footer.gif" valign=middle align="center" colspan="8"><a href="?action=apply" class="cblue">[申请上市]</a><A href="stock.asp" class="cblue">[返回大厅]</A></td></tr>
</table>
<% 
end sub
'-----------------------
sub ChangStat()
	if not master then
		errmess="<li>您没有执行本操作的权限！"
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
					conn.execute("insert into [股票](企业,总股份,剩余股份,交易量,IniTradeNum,开盘价格,当前价格,状态,日期,Explain,LtdImg,TongJi,经营者,Uid)"&_
					     " values('"&rs("Stockname")&"',"&rs("Stocknum")&","&clng(rs("Stocknum")*0.5)&","&temp&","&temp&","&rs("Price")&","&rs("Price")&",'开',date(),'"&replace(rs("Explain"),"'","")&"','"&rs("LtdImg")&"','0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0','"&rs("Username")&"',"&rs("Uid")&")")
					GpID=conn.execute("select top 1 sid from [股票] order by sid desc")(0)
					conn.execute("insert into [大户] (Uid,帐号,sid,买入价格,平均价格,持股数,企业) values ("&rs("uid")&",'"&rs("Username")&"','"&GpID&"',"&rs("Price")&","&rs("Price")&","&clng(rs("Stocknum")*0.5)&",'"&rs("Stockname")&"')")
					PsMoney=rs("Stocknum")*0.5*rs("Price")
					conn.execute("update [客户] set 资金=资金-"&PsMoney*1.02&",总资金=总资金+"&PsMoney*0.98&",持股种类=持股种类+1,今日买入=今日买入+1,总买入=总买入+1 where ID="&rs("uid"))
					call ResetCid()
				elseif cint(rs("States"))=2 then
					conn.execute("update [股票] set [状态]='开' where [企业]='"&rs(0)&"'")
				end if
				conn.execute("update [PersonalStock] set States=1 where sid="&request("sid"))
			end if		
			rs.close
		elseif request("stat")="nopass" then
			set rs=conn.execute("select Stockname,Price,Stocknum,States from [PersonalStock] where sid="&request("sid"))
			if not(rs.eof and rs.bof) then
				if cint(rs(3))=1 then
					conn.execute("update [股票] set [状态]='封' where [企业]='"&rs(0)&"'")
				end if
				conn.execute("update PersonalStock set States=2 where sid="&request("sid"))
			end if
			rs.close		
		end if
		if err.number<>"0" then
			errmess="<li>出现错误："&Err.Description
			call endinfo(1)
			exit sub		
		end if
		response.redirect "?"
	end if	
end sub
'-------------编辑股票---------------
sub Edit()
	dim Sid
	Sid=request("sid")
	if Sid="" then
		errmess="<li>参数错误，请指定相关股票"
		call endinfo(2)
		exit sub	
	end if
	set rs=conn.execute("select * from PersonalStock where sid="&Sid)
	if rs.eof and rs.bof then
		rs.close
		errmess="<li>没有找到指定的上市公司"
		call endinfo(2)
		exit sub
	elseif rs("Uid")<>MyUserID then
		rs.close
		errmess="<li>您不是该上市公司的大股东，没有执行本操作的权限"
		call endinfo(2)
		exit sub		
	end if
%>
<table  width="550" height=300 border=0 cellspacing=1 cellpadding=3 align=center bgcolor="#0066CC">
	<TR>
	<TD BACKGROUND="images/topbg.gif" height=9 colspan="2"></TD>
	</TR>
	<tr>
		<td colspan="2" valign=center align=middle height=23 class=big background="images/header.gif">编辑 个人上市公司 <b><%=rs("Stockname")%></b> 的资料</td>
	</tr>
	<form method=post  action="?action=SaveEdit">
	<input type=hidden name=sid value="<%=rs("sid")%>">
	<tr>
		<td BGCOLOR="#E4E4EF" width="50%"><b>上市公司名称：</b><br>即股票名，最短3Byte，最长14Byte</td>
		<td bgcolor="#FFFFFF" width="50%">&nbsp;<input type=text name=gpname <%if rs("States")<>0 then%> readonly <%end if%> value="<%=rs("Stockname")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF"><b>上市股票单价：</b></td>
		<td bgcolor="#FFFFFF">&nbsp;<%=rs("Price")%> ￥</td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF"><b>上市股票数量：</b></td>
		<td bgcolor="#FFFFFF">&nbsp;<%=rs("Stocknum")%> 股</td>  
	</tr>			
	<tr>
		<td BGCOLOR="#E4E4EF" valign="top"><b>上市公司简介：</b><br>请对对上市公司作简要的介绍，对您的申请的成功率会有影响<br>公司简介不能超过100Byte</td>  
		<td bgcolor="#FFFFFF">&nbsp;<textarea cols="39" name="Explain" rows="5" wrap="PHYSICAL"><%=rs("Explain")%></textarea></td>  
	</tr>
	<tr><td colspan="2" bgcolor="#FFFFFF" align="center"><input type=submit name=submit value=提交><input type=reset name=submit value=重填></td> 
	</form>
	<tr><td height=32 background="images/footer.gif" valign=middle align="center" colspan="8"><A href="stock.asp">[返回股市]</A><A href="<%=Request.ServerVariables("HTTP_REFERER")%>">[返回上一页]</A></td></tr>
</table>
<%	
	rs.close 
end sub 
'-------------保存修改---------------
sub SaveEdit() 
	dim GpName,Explain
	GpName=trim(replace(Request.form("gpname"),"'","")) 
	Explain=trim(replace(Request.form("Explain"),"'","")) 

	if GpName="" then
		errmess="<li>请输入上市公司的名称"
		founderr=true
	elseif len(GpName)>cint(PStock_Setting(3)) or len(GpName)<cint(PStock_Setting(2)) then
		errmess="<li>上市公司名称的长度不能超过 "&PStock_Setting(3)&" Byte或者小于 "&PStock_Setting(2)&" Byte"
		founderr=true
	end if 
	if len(Explain)>cint(PStock_Setting(4)) then
		errmess=errmess+"<li>公司简介不能超过"&PStock_Setting(4)&"Byte"
		founderr=true
	end if		

	if founderr then
		call endinfo(2)
		exit sub	
	end if	
	if Explain="" then Explain="这个家伙很懒，什么都没有留下！"
	conn.execute("update PersonalStock set Stockname='"&GpName&"',Explain='"&Request.form("Explain")&"' where sid="&request.form("sid"))
	sucmess="<li>上市公司信息已经修改成功！"
	rUrl="?"
	call endinfo(2)
end sub
'---------------删除股票-------------
sub del()
	if (not master) and request("reaction")<>"my" then
		errmess="<li>您没有执行本操作的权限！"
		call endinfo(1)
	else
		dim DelID,GPname
		DelID=request("sid")
		if DelID="" then
			errmess="<li>参数错误，请指定要删除的股票"
			call endinfo(2)
			exit sub	
		end if
		set rs=conn.execute("select Stockname,Username,Uid from PersonalStock where sid="&DelID)
		if rs.eof and rs.bof then
			rs.close
			errmess="<li>没有找到你要删除的股票"
			call endinfo(2)
			exit sub
		elseif rs(2)<>MyUserID then
			errmess="<li>非法操作，您没有执行该操作的权限"
			call endinfo(2)
			exit sub
		else			
			GPname=rs(0)
		end if
		rs.close
		conn.execute("delete from PersonalStock where sid="&DelID)
		set rs=conn.execute("select * from [大户] where sid="&DelID)
		
		if request("reaction")<>"my" then
			sucmess="<font color=#66CCFF>新上市公司 "&GPname&"</font><font color=#CCCCCC> 未能获得股票证监会批准，宣布流产</font>"
			conn.execute "insert into RndEvent(content,AddTime) values('"&sucmess&"','"&now()&"' )"  
			sucmess="<li>"&GPname&" 上市公司删除成功"
		else
			sucmess="<li>您已经成功把个股上市的申请注销了"	
		end if
		call endinfo(2)
	end if	
end sub
'-----------------新股上市---------------
sub apply()
	if cint(PStock_Setting(0))=0 then
		errmess="<li>个股上市暂停申请"
		call endinfo(1)
		exit sub		
	end if
	if MyCash<cdbl(PStock_Setting(1)) then
		errmess="<li>您现在的股票帐户资金不足 "&PStock_Setting(1)&" 元，不能申请个人上市公司"
		call endinfo(1)
		exit sub
	end if
	
	dim MinPrice,MaxPrice,YourCash
	set rs=conn.execute("select top 1 开盘价格 from [股票] order by 开盘价格")
	if rs.eof then 
		MinPrice=1.000
	elseif rs(0)*0.5<1 then
		MinPrice=1.000
	else
		MinPrice=formatnumber(rs(0)*0.5,3,-1)
	end if	
	set rs=conn.execute("select top 1 开盘价格 from [股票] order by 开盘价格 desc")
	if rs.eof then 
		MaxPrice=100.000
	elseif rs(0)*0.5<1 then
		MaxPrice=100.000
	else
		MaxPrice=formatnumber(rs(0)*0.5,3,-1)	
	end if	
	YourCash=conn.execute("select 资金 from [客户] where id="&MyUserID&"")(0)
%>
<table  width="80%" border=0 cellspacing=1 cellpadding=3 align=center bgcolor="#0066CC">
	<TR>
	<TD BACKGROUND="images/topbg.gif" height=9 colspan="2"></TD>
	</TR>
	<tr>
		<td colspan="2" valign=center align=middle height=23 class=big background="images/header.gif"><b>个股上市申请表</b></td>
	</tr>
	<form method=post action="?action=SaveApply" name="form1">
	<tr>
		<td BGCOLOR="#E4E4EF" width="50%"><b>申请人：</b></td>
		<td bgcolor="#FFFFFF" width="50%">&nbsp;<%=membername%></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF" width="50%"><b>上市公司名称：</b><br>即股票名，最长<%=PStock_Setting(3)%>Byte，最短<%=PStock_Setting(2)%>Byte</td>
		<td bgcolor="#FFFFFF" width="50%">&nbsp;<input type=text name=gpname value=""> <font color=red>*</font></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF"><b>上市股票单价：</b><br>单价范围：￥<%=MinPrice%>～￥<%=MaxPrice%></td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=price value=""> <font color=red>*</font> <input type=button value='计算最大上市股票数量' name=Button onclick=checknum(<%=MinPrice%>,<%=MaxPrice%>)></td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF"><b>上市股票数量</b><br>每股单价乘以您打算发行的股票数不能大于您现在的账户资金总额</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=gpnum value=""> <font color=red>*</font> <span id="maxnum"></span></td>
	</tr>			
	<tr>
		<td BGCOLOR="#E4E4EF" valign="top"><b>上市公司简介：</b><br>请对对上市公司作简要的介绍，对您的申请的成功率会有影响<br>公司简介不能超过<%=PStock_Setting(4)%>Byte</td>   
		<td bgcolor="#FFFFFF">&nbsp;<textarea cols="39" name="Explain" rows="5" wrap="PHYSICAL">这个家伙很懒，什么都没有留下！</textarea></td> 
	</tr>
	<tr><td colspan="2" bgcolor="#FFFFFF" align="center"><input type=submit name=submit value=申请><input type=reset name=submit value=重填></td>
	<input type=hidden name=MinPrice value="<%=MinPrice%>">
	<input type=hidden name=MaxPrice value="<%=MaxPrice%>">
	<input type=hidden name=YourCash value="<%=YourCash%>">
	<script language="VBScript">
	function checknum(minp,maxp)
		dim price,YourCash
		price=document.form1.price.value
		YourCash=document.form1.YourCash.value
		if price="" then
			msgbox "请输入股票单价",64,"个股上市"
			exit function
		elseif not isnumeric(price) then
			msgbox "股票单价必须输入数值",64,"个股上市"
			exit function
		else
			price=price+0
		end if	
		if price<minp+0 then
			msgbox "股票单价不能小于 "&minp,64,"个股上市"
			exit function
		elseif price>maxp+0 then
			msgbox "股票单价不能大于 "&maxp,64,"个股上市"
			exit function				
		else			
			maxnum.innerHTML="您最多可以上市 <font color=blue>"&int(YourCash/price)&" </font> 股"
		end if
	end function
	</script>
	</form>
	<tr><td height=32 background="images/footer.gif" valign=middle align="center" colspan="8"><A href="stock.asp">[返回股市]</A><A href="<%=Request.ServerVariables("HTTP_REFERER")%>">[返回上一页]</A></td></tr>
</table>
<%
end sub
'-------------保存新股票---------------
sub SaveApply()
	if cint(PStock_Setting(0))=0 then
		errmess="<li>个股上市暂停申请"
		call endinfo(1)
		exit sub		
	end if
	if MyCash<cdbl(PStock_Setting(1)) then
		errmess="<li>您现在的股票帐户资金不足 "&PStock_Setting(1)&" 元，不能申请个人上市公司"
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
	
	if Explain="" then Explain="这个家伙很懒，什么都没有留下！"
	if GpName="" then
		errmess="<li>请输入上市公司的名称"
		founderr=true
	elseif len(GpName)>cint(PStock_Setting(3)) or len(GpName)<cint(PStock_Setting(2)) then
		errmess="<li>上市公司名称的长度不能超过 "&PStock_Setting(3)&" Byte或者小于 "&PStock_Setting(2)&" Byte"
		founderr=true
	end if 
	if len(Explain)>cint(PStock_Setting(4)) then
		errmess=errmess+"<li>公司简介不能超过"&PStock_Setting(4)&"Byte"
		founderr=true
	end if
	if Price="" or (not isnumeric(Price)) then
		errmess=errmess+"<br><li>请输入上市股票单价或者您输入的不是数字"
		founderr=true
	elseif Price+0>MaxPrice+0 then
		errmess=errmess+"<br><li>上市股票单价不能大于 "&MaxPrice&" 元"
		founderr=true
	elseif Price+0<MinPrice+0 then
		errmess=errmess+"<br><li>上市股票单价不能小于 "&MinPrice&" 元"
		founderr=true		
	end if
	if not founderr then	
		MaxGpNum=cdbl(Request.form("YourCash")/Price)
		if gpnum="" or not(isnumeric(gpnum)) then
			errmess=errmess+"<br><li>请输入上市股票数量或者您输入的不是数字"
			founderr=true
		elseif cdbl(gpnum)<cdbl(PStock_Setting(5)) then
			errmess=errmess+"<br><li>股票数量必须大于 "&PStock_Setting(5)&" 股"
			founderr=true		
		elseif cdbl(gpnum)>cdbl(MaxGpNum) then
			errmess=errmess+"<br><li>您最多可以上市的股票数量是 "&MaxGpNum&" 股"
			founderr=true
		end if
	end if	
	
	if founderr then
		call endinfo(2)
		exit sub
	end if
			
	on error resume next 
	set rs=conn.execute("select sid from 股票 where 企业='"&GpName&"'")
	if not rs.eof then
		errmess="<li>您输入的上市公司已经存在，请重新填写上市公司名称"
		endinfo(2)
		exit sub
	end if
	set rs=conn.execute("select Uid from PersonalStock where Stockname='"&GpName&"'")
	if not rs.eof then
		errmess="<li>您输入的上市公司已经上市了，请重新填写上市公司名称"
		endinfo(2)
		exit sub
	end if
	set rs=conn.execute("select Uid from PersonalStock where Uid="&MyUserID&"")	
	if not rs.eof then
		errmess="<li>您已经有自己的上市公司了"
		endinfo(2)
		exit sub		
	end if
			
	conn.execute("insert into PersonalStock(Username,Uid,Stockname,Price,Stocknum,Explain,ApplyDate,States) values('"&membername&"',"&MyUserID&",'"&GpName&"',"&Price&","&gpnum&",'"&Explain&"',now(),0)")
	
	if err.number=0 then
		sucmess="<li>新上市公司 <font color=blue>"&Request.form("gpname")&"</font> 成功申请"
		sucmess=sucmess+"<li>请等待证监会的审批，通过审批之后就会正式上市"
		rUrl="?"
	else
		errmess="<li>出现错误，具体如下："&Err.Description
	end if	
	call endinfo(2)
end sub
'---------------调整股票的Cid--------------------
sub ResetCid()
'说明：Cid 是用于随机事件的，必须是从1开始的自然数列1、2、3、4、5........
	Dim Cid 
    set rs=conn.execute("select sid from [股票] order by sid")
	Cid=1
	do while not rs.eof
		conn.execute("update [股票] set Cid="&Cid&" where sid="&rs(0))
		Cid=Cid+1	
		rs.movenext
	loop
end sub
%>