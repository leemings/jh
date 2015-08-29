<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_gp_conn.asp"-->
<!--#include file="z_gp_Const.asp"-->
<%
stats="个股上市"
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
		<th valign=center align=middle height=25 colspan="8"><b>个人股票上市申请一览表</b> | <a href="?action=apply" class="cblue">申请个股上市</a></th>
	</tr>
	<tr align="center" height=20 valign="middle"> 
		<td class=tablebody2 align=center>上市公司</td>
		<td class=tablebody2 align=center>股票数量</td>
		<td class=tablebody2 align=center>上市价格</td>
		<td class=tablebody2 align=center>申请人</td>
		<td class=tablebody2 align=center>申请日期</td>
		<td class=tablebody2 align=center>公司简介</td>
		<td class=tablebody2 align=center>审批</td>
		<td class=tablebody2 align=center>操作</td>
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
		response.write "<tr><td class=tablebody1 colspan=8>暂时没有人申请个股上市股票</td></tr>"
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
				<td class=tablebody1><%if rs("States")=0 then%>未审核<%elseif rs("States")=1 then%>已批准<%else%>不批准<%end if%></td>
				<td class=tablebody1><%
					if master then
						if rs("States")=0 then
							%><a href="?action=ChangStat&stat=pass&sid=<%=rs("sid")%>">批准</a> | <a href="?action=ChangStat&stat=nopass&sid=<%=rs("sid")%>">不批准</a><%elseif rs("States")=1 then%><a href="?action=ChangStat&stat=nopass&sid=<%=rs("sid")%>" class=cred>摘牌</a><%else%><a href="?action=ChangStat&stat=pass&sid=<%=rs("sid")%>">批准</a><%end if%> | <a href="?action=edit&sid=<%=rs("sid")%>">编辑</a> | <a href="?action=del&sid=<%=rs("sid")%>" onclick="javascript:{if(confirm('您确定删除 <%=rs("Stockname")%> 这个股票吗?')){return true;}return false;}" class=cred>删除</a><%
						elseif rs("Uid")=MyUserID then
							%><a href="?action=edit&sid=<%=rs("sid")%>">编辑</a> <%if rs("States")=0 then%>| <a href="?action=del&reaction=my&sid=<%=rs("sid")%>" onclick="javascript:{if(confirm('您确定取消 <%=rs("Stockname")%> 这个股票的上市申请吗?')){return true;}return false;}" class=cred>注销</a><%
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
			<td class=tablebody2 colspan=2>共有 <font color=blue><%=totalrec%></font> 个申请，每页 <font color=blue><%=rs.PageSize%></font> 个，第 <font color=red><%=currentpage%></font> 页 / 共 <font color=blue><%=Pcount%></font> 页</td>
			<td class=tablebody2 align=right colspan=6>分页：<%call disppagenum(currentpage,pcount,"""?page=","""")%>&nbsp;&nbsp;&nbsp;&nbsp;</td>
		</tr>
	<%end if
	rs.close
	set rs=nothing%>			
	<tr>
		<th height=25 valign=middle align="center" colspan="8"><a href="?action=apply" class="cblue">[申请上市]</a><A href="z_gp_Gupiao.asp" class="cblue">[返回大厅]</A></th>
	</tr>
<%end sub
'-----------------------
sub ChangStat()
	if not master then
		errmess="<li>您没有执行本操作的权限！"
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
					     " values('"&rs("Stockname")&"',"&rs("Stocknum")&","&clng(rs("Stocknum")*0.5)&","&temp&","&rs("Price")&","&rs("Price")&",'开',date(),'"&replace(rs("Explain"),"'","")&"|"&temp&"','"&rs("LtdImg")&"','0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0','"&rs("Username")&"',"&rs("Uid")&")")
					GpID=gp_conn.execute("select top 1 sid from [GuPiao] order by sid desc")(0)
					gp_conn.execute("insert into [DaHu] (Uid,ZhangHao,sid,MaiRuJiaGe,PingJunJiaGe,ChiGuShu,QiYe) values ("&rs("uid")&",'"&rs("Username")&"','"&GpID&"',"&rs("Price")&","&rs("Price")&","&clng(rs("Stocknum")*0.5)&",'"&rs("Stockname")&"')")
					PsMoney=rs("Stocknum")*0.5*rs("Price")
					gp_conn.execute("update [KeHu] set ZiJin=ZiJin-"&PsMoney*1.02&",ZongZiJin=ZongZiJin+"&PsMoney*0.98&",ChiGuZhongLei=ChiGuZhongLei+1,JinRiMaiRu=JinRiMaiRu+1,ZongMaiRu=ZongMaiRu+1 where ID="&rs("uid"))
					call ResetCid()
				elseif CLNG(rs("States"))=2 then
					gp_conn.execute("update [GuPiao] set [ZhuangTai]='开' where [QiYe]='"&rs(0)&"'")
				end if
				gp_conn.execute("update [PersonalStock] set States=1 where sid="&request("sid"))
			end if		
			rs.close
		elseif request("stat")="nopass" then
			set rs=gp_conn.execute("select Stockname,Price,Stocknum,States from [PersonalStock] where sid="&request("sid"))
			if not(rs.eof and rs.bof) then
				if CLNG(rs(3))=1 then
					gp_conn.execute("update [GuPiao] set [ZhuangTai]='封' where [QiYe]='"&rs(0)&"'")
				end if
				gp_conn.execute("update PersonalStock set States=2 where sid="&request("sid"))
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
	set rs=gp_conn.execute("select * from PersonalStock where sid="&Sid)
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
	end if%>
	<tr>
		<th colspan="2" valign=center align=middle height=25>编辑 个人上市公司 <b><%=rs("Stockname")%></b> 的资料</th>
	</tr>
	<form method=post  action="?action=SaveEdit">
	<input type=hidden name=sid value="<%=rs("sid")%>">
	<tr>
		<td class=tablebody1 width="50%"><b>上市公司名称：</b><br>即股票名，最短3Byte，最长14Byte</td>
		<td class=tablebody1 width="50%">&nbsp;<input type=text name=gpname <%if rs("States")<>0 then%> readonly <%end if%> value="<%=rs("Stockname")%>"></td>
	</tr>
	<tr>
		<td class=tablebody1><b>上市股票单价：</b></td>
		<td class=tablebody1>&nbsp;<%=rs("Price")%></td>
	</tr>	
	<tr>
		<td class=tablebody1><b>上市股票数量：</b></td>
		<td class=tablebody1>&nbsp;<%=rs("Stocknum")%> 股</td>  
	</tr>			
	<tr>
		<td class=tablebody1 valign="top"><b>上市公司简介：</b><br>请对对上市公司作简要的介绍，对您的申请的成功率会有影响<br>公司简介不能超过100Byte</td>  
		<td class=tablebody1>&nbsp;<textarea cols="39" name="Explain" rows="5" wrap="PHYSICAL"><%=rs("Explain")%></textarea></td>  
	</tr>
	<tr>
		<td class=tablebody2 colspan="2" align="center"><input type=submit name=submit value=提交><input type=reset name=submit value=重填></td> 
	</tr>
	</form>
	<tr>
		<th height=25 valign=middle align="center" colspan="2"><A href="z_gp_Gupiao.asp">[返回股市]</A><A href="<%=Request.ServerVariables("HTTP_REFERER")%>">[返回上一页]</A></th>
	</tr>
	<%rs.close 
end sub 
'-------------保存修改---------------
sub SaveEdit() 
	dim GpName,Explain
	GpName=trim(replace(Request.form("gpname"),"'","")) 
	Explain=trim(replace(Request.form("Explain"),"'","")) 

	if GpName="" then
		errmess="<li>请输入上市公司的名称"
		founderr=true
	elseif len(GpName)>PStock_Setting(3) or len(GpName)<PStock_Setting(2) then
		errmess="<li>上市公司名称的长度不能超过 "&PStock_Setting(3)&" Byte或者小于 "&PStock_Setting(2)&" Byte"
		founderr=true
	end if 
	if len(Explain)>PStock_Setting(4) then
		errmess=errmess+"<li>公司简介不能超过"&PStock_Setting(4)&"Byte"
		founderr=true
	end if		

	if founderr then
		call endinfo(2)
		exit sub	
	end if	
	if Explain="" then Explain="这个家伙很懒，什么都没有留下！"
	gp_conn.execute("update PersonalStock set Stockname='"&GpName&"',Explain='"&Request.form("Explain")&"' where sid="&request.form("sid"))
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
		set rs=gp_conn.execute("select Stockname,Username,Uid from PersonalStock where sid="&DelID)
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
		gp_conn.execute("delete from PersonalStock where sid="&DelID)
		set rs=gp_conn.execute("select * from [DaHu] where sid="&DelID)
		
		if request("reaction")<>"my" then
			sucmess="<font color=#66CCFF>新上市公司 "&GPname&"</font><font color=#CCCCCC> 未能获得股票证监会批准，宣布流产</font>"
			gp_conn.execute "insert into RndEvent(content,AddTime) values('"&sucmess&"','"&now()&"' )"  
			sucmess="<li>"&GPname&" 上市公司删除成功"
		else
			sucmess="<li>您已经成功把个股上市的申请注销了"	
		end if
		call endinfo(2)
	end if	
end sub
'-----------------新股上市---------------
sub apply()
	if PStock_Setting(0)=0 then
		errmess="<li>个股上市暂停申请"
		call endinfo(1)
		exit sub		
	end if
	if MyCash<csng(PStock_Setting(1)) then
		errmess="<li>您现在的股票帐户资金不足 "&PStock_Setting(1)&" 元，不能申请个人上市公司"
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
		<th colspan="2" valign=center align=middle height=25><b>个股上市申请表</b></th>
	</tr>
	<form method=post action="?action=SaveApply" name="form1">
	<tr>
		<td class=tablebody1 width="50%"><b>申请人：</b></td>
		<td class=tablebody1 width="50%">&nbsp;<%=membername%></td>
	</tr>
	<tr>
		<td class=tablebody1 width="50%"><b>上市公司名称：</b><br>即股票名，最长<%=PStock_Setting(3)%>Byte，最短<%=PStock_Setting(2)%>Byte</td>
		<td class=tablebody1 width="50%">&nbsp;<input type=text name=gpname value=""> <font color=red>*</font></td>
	</tr>
	<tr>
		<td class=tablebody1><b>上市股票单价：</b><br>单价范围：<%=MinPrice%>～<%=MaxPrice%></td>
		<td class=tablebody1>&nbsp;<input type=text name=price value=""> <font color=red>*</font> <input type=button value='计算最大上市股票数量' name=Button onclick=checknum(<%=MinPrice%>,<%=MaxPrice%>)></td>
	</tr>	
	<tr>
		<td class=tablebody1><b>上市股票数量</b><br>每股单价乘以您打算发行的股票数不能大于您现在的账户资金总额</td>
		<td class=tablebody1>&nbsp;<input type=text name=gpnum value=""> <font color=red>*</font> <span id="maxnum"></span></td>
	</tr>			
	<tr>
		<td class=tablebody1 valign="top"><b>上市公司简介：</b><br>请对上市公司作简要的介绍，对您的申请的成功率会有影响<br>公司简介不能超过<%=PStock_Setting(4)%>Byte</td>   
		<td class=tablebody1>&nbsp;<textarea cols="39" name="Explain" rows="5" wrap="PHYSICAL">这个家伙很懒，什么都没有留下！</textarea></td> 
	</tr>
	<tr>
		<td class=tablebody2 colspan="2" align="center"><input type=submit name=submit value=申请><input type=reset name=submit value=重填></td>
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
	<tr>
		<th height=25 valign=middle align="center" colspan="2"><A href="z_gp_Gupiao.asp">[返回股市]</A><A href="<%=Request.ServerVariables("HTTP_REFERER")%>">[返回上一页]</A></th>
	</tr>
<%end sub
'-------------保存新股票---------------
sub SaveApply()
	if PStock_Setting(0)=0 then
		errmess="<li>个股上市暂停申请"
		call endinfo(1)
		exit sub		
	end if
	if MyCash<csng(PStock_Setting(1)) then
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
	elseif len(GpName)>PStock_Setting(3) or len(GpName)<PStock_Setting(2) then
		errmess="<li>上市公司名称的长度不能超过 "&PStock_Setting(3)&" Byte或者小于 "&PStock_Setting(2)&" Byte"
		founderr=true
	end if 
	if len(Explain)>PStock_Setting(4) then
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
		MaxGpNum=clng(Request.form("YourCash")/Price)
		if gpnum="" or not(isnumeric(gpnum)) then
			errmess=errmess+"<br><li>请输入上市股票数量或者您输入的不是数字"
			founderr=true
		elseif clng(gpnum)<clng(PStock_Setting(5)) then
			errmess=errmess+"<br><li>股票数量必须大于 "&PStock_Setting(5)&" 股"
			founderr=true		
		elseif clng(gpnum)>clng(MaxGpNum) then
			errmess=errmess+"<br><li>您最多可以上市的股票数量是 "&MaxGpNum&" 股"
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
		errmess="<li>您输入的上市公司已经存在，请重新填写上市公司名称"
		endinfo(2)
		exit sub
	end if
	set rs=gp_conn.execute("select Uid from PersonalStock where Stockname='"&GpName&"'")
	if not rs.eof then
		errmess="<li>您输入的上市公司已经上市了，请重新填写上市公司名称"
		endinfo(2)
		exit sub
	end if
	set rs=gp_conn.execute("select Uid from PersonalStock where Uid="&MyUserID&"")	
	if not rs.eof then
		errmess="<li>您已经有自己的上市公司了"
		endinfo(2)
		exit sub		
	end if
			
	gp_conn.execute("insert into PersonalStock(Username,Uid,Stockname,Price,Stocknum,Explain,ApplyDate,States) values('"&membername&"',"&MyUserID&",'"&GpName&"',"&Price&","&gpnum&",'"&Explain&"',now(),0)")
	
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
    set rs=gp_conn.execute("select sid from [GuPiao] order by sid")
	Cid=1
	do while not rs.eof
		gp_conn.execute("update [GuPiao] set Cid="&Cid&" where sid="&rs(0))
		Cid=Cid+1	
		rs.movenext
	loop
end sub
%>