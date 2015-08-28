<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html><head><title><%=Gupiao_Setting(5)%>-股市控制面板</title>
<!--#include file="css.asp"-->
</head><body bgcolor="#ffffff" text="#000000" style="FONT-SIZE: 9pt" topmargin=5 leftmargin=0 oncontextmenu=self.event.returnValue=false>
<table  width="98%" border=0 cellspacing=1 cellpadding=0 align=center bgcolor="#0066CC">
	<TR>
		<TD BACKGROUND="Images/topbg.gif" height=9 colspan=3>
	</TD>
	</TR>
	<tr>
		<td valign=center align=middle height=23 background="Images/Header.gif"><font size="2"><b>股市管理中心</b></font></td>
	</tr>
	<tr><td bgcolor="#E4E8EF"> 
<% 
call AdminHead()
if session("yx8_mhjh_username")<>Application("yx8_mhjh_admin") then
	errmess="<li>本页为管理员专用，您没有管理本页的权限！"
	call endinfo(1)
else 
	select case request("action")
		case "search"
			call search()
		case "edit"
			call Edit()
		case "SaveEdit"		
			call SaveEdit()
		case "del"
			call del()
		case "newgupiao"
			call newgupiao()
		case "SaveNew"		
			call SaveNew()
		case "ChangStat"
			call ChangStat()					
		case else
			call main()
	end select
end if
%>
<br>
</td></tr>
<tr><td height=32 background="images/footer.gif" valign=middle></td></tr>
</table>
</body></html>
<%
CloseDatabase		'关闭数据库
'=====================================
sub main()
%>
<table width="97%" border="0" cellspacing="1" cellpadding="3" align="center" bgcolor="#0066CC">
	<tr>
		<td valign=center align=middle height=23 background="Images/Header.gif" colspan="8">股票管理 | <a href="?action=newgupiao" class="cblue">新股上市</a></td>
	</tr>
	<tr> 
		<form name="form1" method="post" action="?action=search">
		<td colspan="8" height=20 bgcolor="#ffffff">快速查找：<input type="text"  name=gname ><input type="checkbox"  name="usernamechk" value="yes" checked>名称完全匹配<input type="submit" name="Submit" value="查询股票">
		<div align="right"><font color=red>[注意] 删除操作直接删除相应的股票，持股者会以当前价格抛出所有股票</font></div>
		</td>
		</form>
	</tr>
	<tr align="center" > 
		<td align="center" background="Images/lan.gif"><font color=#ffffff>代号</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>股票名称</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>总股份</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>昨日收盘价</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>目前成交价</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>成交量</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>状态</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>操作</font></td>
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
	sql= "select * from 股票"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "<tr><td colspan=8 bgcolor=#E4E8EF>还没有上市股票啊</td></tr>"
	else
		rs.PageSize = 20 
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		do while (not rs.eof) and (not page_count = 20)	
%>
  <tr align="center"> 
    <td bgcolor="#FFFFFF" ><a href="exchange.asp?sid=<%=rs("sid")%>"> <%=rs("sid")%> </a></td>
    <td bgcolor="#FFFFFF"><a href="?action=edit&sid=<%=rs("sid")%>"><%=rs("企业")%></a></td>
    <td bgcolor="#FFFFFF"><%=formatnumber(rs("总股份"),0)%></td>
    <td bgcolor="#FFFFFF"><%=formatcurrency(rs("开盘价格"),3)%></td>
    <td bgcolor="#FFFFFF"><%=formatcurrency(rs("当前价格"),3)%></td>
	<td bgcolor="#FFFFFF"><%=rs("成交量")%></td>
	<td bgcolor="#FFFFFF"><%if rs("状态")="封" then%><font color=red><%end if%><%=rs("状态")%></td>
    <td bgcolor="#FFFFFF"><%if rs("状态")="开" then%><a href="?action=ChangStat&stat=closethis&sid=<%=rs("sid")%>">封盘</a><%else%><a href="?action=ChangStat&stat=openthis&sid=<%=rs("sid")%>" class=cred>开市</a><%end if%> | <a href="?action=edit&sid=<%=rs("sid")%>">编辑</a> | <a href="?action=del&sid=<%=rs("sid")%>" onclick="javascript:{if(confirm('您确定删除 <%=rs("企业")%> 这个股票吗?')){return true;}return false;}" class=cred>删除</a></td>
  </tr>
<%         
		page_count = page_count + 1		
		rs.movenext
	loop
	Pcount=rs.PageCount         
%>
	<tr><td colspan=8 background="images/title.gif">
		<table border="0" cellpadding="0" cellspacing="0" width="100%"><tr>
			<td align=left>共有<font color=blue><%=totalrec%></font>支股票，每页<font color=blue><%=rs.PageSize%></font>支，第<font color=red><%=currentpage%></font>页/共<font color=blue><%=Pcount%></font>页</td>
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
				
		end if
		rs.close
		set rs=nothing	
%>			
			</td></tr>
		 </table>
	</td></tr>	  	
</table>
<% 
end sub
'-----------------------
sub ChangStat()
	on error resume next
	if request("stat")="closethis" then
		conn.execute("update 股票 set 状态='封' where sid="&request("sid"))
	elseif request("stat")="openthis" then
		conn.execute("update 股票 set 状态='开' where sid="&request("sid"))		
	end if
	response.redirect "?"
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
	set rs=conn.execute("select * from [股票] where sid="&Sid)
	if rs.eof and rs.bof then
		rs.close
		errmess="<li>没有找到该股票"
		call endinfo(2)
		exit sub
	end if
%>
<table  width="75%" border=0 cellspacing=1 cellpadding=3 align=center bgcolor="#0066CC">
	<TR>
	<TD BACKGROUND="images/topbg.gif" height=9 colspan="2"></TD>
	</TR>
	<tr>
		<td colspan="2" valign=center align=middle height=23 class=big background="images/header.gif">编辑 <font face="Courier New, Courier, mono" color="#FF0000"><%=sid%></font> 号股票<%=rs("企业")%></td>
	</tr>
	<form method=post  action="?action=SaveEdit">
	<tr>
		<td BGCOLOR="#E4E4EF" width="40%">代号</td>
		<td bgcolor="#FFFFFF" width="60%">&nbsp;<%=rs("sid")%><input type=hidden name=sid value="<%=rs("sid")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF">经营者</td>
		<td bgcolor="#FFFFFF">&nbsp;<%=rs("经营者")%></td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF">企业</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=gpname value="<%=rs("企业")%>"><input type=hidden name=gpoldname value="<%=rs("企业")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF">总股份</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=TotalStockNum value="<%=rs("总股份")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF">初始交易量</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=inigpnum value="<%=rs("IniTradeNum")%>"></td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF">剩余交易量</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=gpnum value="<%=rs("交易量")%>"></td>
	</tr>			
	<tr>
		<td BGCOLOR="#E4E4EF">开盘价格</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=openprice value="<%=rs("开盘价格")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF">当前价格</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=nowprice value="<%=rs("当前价格")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF">上市日期</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=gpdate value="<%=rs("日期")%>"></td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF">状态</td>
		<td bgcolor="#FFFFFF"><input type=radio name=state value="开" <%if rs("状态")="开" then%> checked <%end if%>> 开<input type=radio name=state value="封" <%if rs("状态")="封" then%> checked <%end if%>>封</td>
	</tr>

	<tr>
		<td BGCOLOR="#E4E4EF">公司图片(150*190)<br>如果图片路径是相对路径，则当前目录是股票的根目录<br>图片也可以是网上的图片连接</td> 
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=LtdImg size="39" value="<%=rs("LtdImg")%>"></td> 
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF" valign="top"><b>公司简介：</b><br>上市公司简绍<br>公司简介不能超过100Byte</td>   
		<td bgcolor="#FFFFFF">&nbsp;<textarea cols="39" name="Explain" rows="5" wrap="PHYSICAL"><%=htmlencode(rs("Explain"))%></textarea></td>  
	</tr>

	<tr>
		<td BGCOLOR="#E4E4EF"><font color=gray>剩余股份</font></td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text readonly name=remaingp value="<%=rs("剩余股份")%>"></td> 
	</tr>			
	<tr>
		<td BGCOLOR="#E4E4EF"><font color=gray>今日成交量</font></td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text readonly name=chengjiao value="<%=rs("成交量")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF"><font color=gray>今日买入笔数</font></td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text readonly name=buy value="<%=rs("买入笔数")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF"><font color=gray>今日卖出笔数</font></td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text readonly name=sale value="<%=rs("卖出笔数")%>"></td>
	</tr>					

	<tr><td colspan="2" bgcolor="#FFFFFF" align="center"><input type=submit name=submit value=执行修改></td>
	</form>
</table>
<%	
	rs.close
end sub
'-------------保存修改---------------
sub SaveEdit()
	dim GpName,gpnum,OpenPrice,NowPrice,inigpnum
	dim TotalStockNum,gpdate,LtdImg,Explain,gpoldname
	GpName=trim(replace(Request.form("gpname"),"'","")) 
	gpoldname=trim(replace(Request.form("gpoldname"),"'","")) 
	gpnum=trim(Request.form("gpnum"))
	inigpnum=trim(Request.form("inigpnum"))
	
	OpenPrice=trim(Request.form("openprice"))
	NowPrice=trim(Request.form("nowprice"))
	
	TotalStockNum=trim(Request.form("TotalStockNum")) 
	gpdate=trim(Request.form("gpdate"))
	LtdImg=trim(Request.form("LtdImg"))
	Explain=trim(replace(Request.form("Explain"),"'",""))	
	
	if GpName="" then
		errmess="<li>请输入上市股票的名称"
		call endinfo(2)
		exit sub
	end if	
	errmess=""
	if not isnumeric(gpnum) then 
		errmess=errmess+"<li>交易量必须输入数字"
	end if
	if not isnumeric(inigpnum) then 
		errmess=errmess+"<li>初始交易量必须输入数字"
	end if		
	if not isnumeric(OpenPrice) then 
		errmess=errmess+"<li>开盘价格必须输入数字"
	end if	
	if not isnumeric(NowPrice) then 
		errmess=errmess+"<li>当前价格必须输入数字"
	end if
	
	if not isnumeric(TotalStockNum) then 
		errmess=errmess+"<li>股票总数必须是数字"
	end if
	
	if gpoldname<>GpName then
		set rs=conn.execute("select sid from [股票] where [企业]='"&GpName&"'")
		if not rs.eof then
			errmess="<li>您输入的上市公司已经存在，请重新填写上市公司名称"
		end if
	end if	
	
	if errmess<>"" then
		call endinfo(2)
		exit sub	
	end if	
	if isdate(gpdate) then 
		gpdate=" ,日期='"&gpdate&"'"
	else
		gpdate=""
	end if
	if Explain="" then Explain="暂无"
			
	conn.execute("update [股票] set [企业]='"&GpName&"',交易量="&gpnum&",IniTradeNum="&inigpnum&",开盘价格="&OpenPrice&",当前价格="&NowPrice&",状态='"&Request.form("state")&"',总股份="&TotalStockNum&",LtdImg='"&LtdImg&"',Explain='"&Explain&"' "&gpdate&" where sid="&request("sid"))
	if gpoldname<>GpName then	'改变股票名字时 request("sid")
		conn.execute("update [大户] set 企业='"&GpName&"' where sid="&request("sid")) 
		conn.execute("update [PersonalStock] set Stockname='"&GpName&"' where Stockname='"&gpoldname&"'")
	end if
	
	sucmess="<li>股票信息已经修改完毕!"
	rUrl="?"
	call endinfo(2)
end sub
'---------------删除股票-------------
sub del()
	dim DelID,AddMoney,NowPrice,trs,GPname
	DelID=request("sid")
	if DelID="" then
		errmess="参数错误，请指定要删除的股票"
		call endinfo(2)
		exit sub	
	end if
	set rs=conn.execute("select 当前价格,企业 from 股票 where sid="&DelID)
	if rs.eof and rs.bof then
		rs.close
		errmess="没有找到你要删除的股票"
		call endinfo(2)
		exit sub
	else
		NowPrice=rs(0)
		GPname=rs(1)
	end if
	rs.close
	conn.execute("delete from [股票] where sid="&DelID)
	set rs=conn.execute("select * from [大户] where sid="&DelID)
	do while not rs.eof
		AddMoney=rs("持股数")*NowPrice
		conn.execute "update [客户] set 资金=资金+"&AddMoney&",总资金=总资金-"&AddMoney&",今日卖出=今日卖出+1,总卖出=总卖出+1 where id="&rs("uid")&""
		rs.movenext
	loop
	rs.close
	conn.execute("delete from [大户] where sid="&DelID)
	sucmess="<font color=#66CCFF>"&GPname&"</font><font color=#CCCCCC>倒闭，购买了该股的客户被迫以当前价格全部抛出</font>"
	conn.execute "insert into RndEvent(content,addtime) values('"&sucmess&"','"&now()&"' )"  
	call ResetCid()		'重新排列股票的Cid
	sucmess="<li>"&GPname&" 股票删除成功，购买了该股票的所有客户被强迫以当前价格抛出"
	call endinfo(2)
end sub
'---------------搜索股票-------------
sub search()
	dim GupiaoName
	GupiaoName=trim(replace(request("gname"),"'",""))
	if GupiaoName="" then
		errmess="<li>请输入查找关键字"
		call endinfo(2)
		exit sub
	end if	
	if request("usernamechk")="yes" then
		sql="select * from 股票 where 企业 = '"&GupiaoName&"'"  
	else	
		sql="select * from 股票 where 企业 like '%"&GupiaoName&"%'"         
	end if	
	
%>
<table border="0" width="97%" bgcolor="#0066CC" cellspacing="1" cellpadding="3" align="center">
	<tr align="center" > 
		<td align="center" background="Images/lan.gif"><font color=#ffffff>代号</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>股票名称</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>总股份</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>昨日收盘价</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>目前成交价</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>成交量</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>状态</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>操作</font></td>
	</tr>
<%		
	set rs=conn.execute(sql)         
	if rs.eof then
		response.Write "<tr><td bgcolor=""#FFFFFF"" colspan=8>没有找到任何股票</td></tr>"
	else
		do while not rs.eof
%>
	<tr align="center"> 
		<td bgcolor="#FFFFFF" ><a href="exchange.asp?sid=<%=rs("sid")%>"> <%=rs("sid")%> </a></td>
		<td bgcolor="#FFFFFF"><a href="?action=edit&sid=<%=rs("sid")%>"><%=rs("企业")%></a></td>
		<td bgcolor="#FFFFFF"><%=formatnumber(rs("总股份"),0)%></td>
		<td bgcolor="#FFFFFF"><%=formatcurrency(rs("开盘价格"),3)%></td>
		<td bgcolor="#FFFFFF"><%=formatcurrency(rs("当前价格"),3)%></td>
		<td bgcolor="#FFFFFF"><%=rs("成交量")%></td>
		<td bgcolor="#FFFFFF"><%if rs("状态")="封" then%><font color=red><%end if%><%=rs("状态")%></td>
		<td bgcolor="#FFFFFF"><a href="?action=edit&sid=<%=rs("sid")%>">编辑</a> | <a href="?action=del&sid=<%=rs("sid")%>" onclick="javascript:{if(confirm('您确定删除 <%=rs("企业")%> 这个股票吗?')){return true;}return false;}">删除</a></td>
	</tr>	
  <%
			rs.movenext
		loop
		rs.close
	end if
%>
	<TR><TD BACKGROUND="Images/title.gif" height=21 align="center" colspan="8"><A href="?">[返回列表]</A></TD></TR>
</table>
<%	
end sub
'-----------------新股上市---------------
sub newgupiao()
%>
<table  width="75%" border=0 cellspacing=1 cellpadding=3 align=center bgcolor="#0066CC">
	<TR>
	<TD BACKGROUND="images/topbg.gif" height=9 colspan="2"></TD>
	</TR>
	<tr>
		<td colspan="2" valign=center align=middle height=23 class=big background="images/header.gif">新股上市</td>
	</tr>
	<form method=post  action="?action=SaveNew">
	<tr>
		<td BGCOLOR="#E4E4EF" width="40%">ID</td>
		<td bgcolor="#FFFFFF" width="60%">&nbsp;自动分配</td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF"><b>企业(股票名)</b>字数不要超过20字</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=gpname value=""> <font color=red>*</font></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF"><b>交易量</b><br>本日交易量</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=gpnum value="10000"> 股</td> 
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF"><b>总股份</b><br>股票总数量</td> 
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=TotalStockNum value="10000000"> 股</td>
	</tr>				
	<tr>
		<td BGCOLOR="#E4E4EF">开盘价格</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=openprice value="10.00"> 元</td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF">当前价格</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=nowprice value="10.00"> 元</td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF">状态</td>
		<td bgcolor="#FFFFFF"><input type=radio name=state value="开" checked> 开<input type=radio name=state value="封">封</td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF"><b>公司图片(150*190)</b><br>如果图片路径是相对路径，则当前目录是股票的根目录<br>图片也可以是网上的图片连接</td> 
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=LtdImg size="39" value="Images/LtdImg/1.jpg"></td> 
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF" valign="top"><b>公司简介：</b><br>上市公司简绍<br>公司简介不能超过100Byte</td>   
		<td bgcolor="#FFFFFF">&nbsp;<textarea cols="39" name="Explain" rows="5" wrap="PHYSICAL">这个家伙很懒，什么都没有留下！</textarea></td>  
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF">成交量</td>
		<td bgcolor="#FFFFFF">&nbsp;0</td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF">涨跌</td>
		<td bgcolor="#FFFFFF">&nbsp;0</td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF">买入笔数</td>
		<td bgcolor="#FFFFFF">&nbsp;0</td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF">卖出笔数</td>
		<td bgcolor="#FFFFFF">&nbsp;0</td>
	</tr>					

	<tr><td colspan="2" bgcolor="#FFFFFF" align="center"><input type=submit name=submit value=增加><input type=reset name=submit value=重填></td>
	</form>
</table>
<%
end sub
'-------------保存新股票---------------
sub SaveNew()
	dim GpName,gpnum,OpenPrice,NowPrice
	dim TotalStockNum,LtdImg,Explain,gpoldname
	GpName=trim(replace(Request.form("gpname"),"'","")) 
	gpnum=trim(Request.form("gpnum"))
	OpenPrice=trim(Request.form("openprice"))
	NowPrice=trim(Request.form("nowprice"))
	
	if GpName="" then
		errmess="<li>请输入上市股票的名称"
		call endinfo(2)
		exit sub
	end if	
	TotalStockNum=trim(Request.form("TotalStockNum")) 
	LtdImg=trim(Request.form("LtdImg"))
	Explain=trim(replace(Request.form("Explain"),"'",""))	
	
	if GpName="" then
		errmess="<li>请输入上市股票的名称"
		call endinfo(2)
		exit sub
	end if	
	errmess=""
	if not isnumeric(gpnum) then 
		errmess=errmess+"<li>剩余交易量必须输入数字"
	end if
	if not isnumeric(OpenPrice) then 
		errmess=errmess+"<li>开盘价格必须输入数字"
	end if	
	if not isnumeric(NowPrice) then 
		errmess=errmess+"<li>当前价格必须输入数字"
	end if
	
	if not isnumeric(TotalStockNum) then 
		errmess=errmess+"<li>股票总数必须是数字"
	end if
	if errmess<>"" then
		call endinfo(2)
		exit sub	
	end if	
	if Explain="" then Explain="暂无"
	on error resume next 
	set rs=conn.execute("select sid from 股票 where 企业='"&GpName&"'")
	if not rs.eof then
		errmess="<li>您输入的股票名称已经上市了，请重新填写股票名称"
		endinfo(2)
		exit sub
	else
		conn.execute("insert into 股票(企业,交易量,开盘价格,当前价格,状态,日期,总股份,剩余股份,LtdImg,Explain,IniTradeNum,TongJi) values('"&GpName&"',"&gpnum&","&OpenPrice&","&NowPrice&",'"&Request.form("state")&"',date(),"&TotalStockNum&","&TotalStockNum&",'"&LtdImg&"','"&Explain&"',"&gpnum&",'0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0')")
	end if
	if err.number=0 then
		sucmess="<li>恭喜，新股票 <font color=blue>"&Request.form("gpname")&"</font> 成功上市"
		rUrl="?"
		call ResetCid()
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