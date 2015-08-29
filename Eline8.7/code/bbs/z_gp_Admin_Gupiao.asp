<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_gp_conn.asp"-->
<!--#include file="z_gp_Const.asp"-->
<%
stats="股票管理"
call nav()
call head_var(0,0,GuPiao_Setting(5),"z_gp_gupiao.asp")
if not master then
	Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	call dvbbs_error()
else
	call AdminHead()%>
	<table class=tableborder1 cellspacing=1 cellpadding=3 align=center border=0 width="<%=Forum_body(12)%>">
	<%select case request("action")
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
	end select%>
	</table>
<%end if
call footer()
'=====================================
sub main()%>
	<tr>
		<th valign=center align=middle height=25 colspan="8">股票管理 | <a href="?action=newgupiao" class="cblue">新股上市</a></th>
	</tr>
	<tr> 
		<form name="form1" method="post" action="?action=search">
		<td class=tablebody2 colspan="8" height=20>快速查找：<input type="text"  name=gname ><input type="checkbox"  name="usernamechk" value="yes" checked>名称完全匹配<input type="submit" name="Submit" value="查询股票">
		<div align="right"><font color=red>[注意] 删除操作直接删除相应的股票，持股者会以当前价格抛出所有股票</font></div>
		</td>
		</form>
	</tr>
	<tr align="center" > 
		<th align="center">代号</th>
		<th align="center">股票名称</th>
		<th align="center">总股份</th>
		<th align="center">昨日收盘价</th>
		<th align="center">目前成交价</th>
		<th align="center">成交量</th>
		<th align="center">状态</th>
		<th align="center">操作</th>
	</tr>
	<%dim currentpage,page_count,Pcount
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
	sql= "select * from [GuPiao]"
	rs.open sql,gp_conn,1,1
	if rs.eof and rs.bof then
		response.write "<tr><td class=tablebody1 colspan=8>还没有上市股票啊</td></tr>"
	else
		rs.PageSize = Gupiao_Setting(2) 
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		do while (not rs.eof) and (not page_count = rs.PageSize)%>
		  <tr align="center"> 
		    <td class=tablebody1><a href="z_gp_Exchange.asp?sid=<%=rs("sid")%>"> <%=rs("sid")%> </a></td>
		    <td class=tablebody1><a href="?action=edit&sid=<%=rs("sid")%>"><%=rs("QiYe")%></a></td>
		    <td class=tablebody1 align=right><%=formatnumber(rs("JiaoYiLiang"),0)%>&nbsp;</td>
		    <td class=tablebody1 align=right><%=formatnumber(rs("KaiPanJiaGe"),2)%>&nbsp;</td>
		    <td class=tablebody1 align=right><%=formatnumber(rs("DangQianJiaGe"),2)%>&nbsp;</td>
				<td class=tablebody1 align=right><%=rs("ChengJiaoLiang")%>&nbsp;</td>
				<td class=tablebody1><%if rs("ZhuangTai")="封" then%><font color=red><%end if%><%=rs("ZhuangTai")%></td>
		    <td class=tablebody1><%if rs("ZhuangTai")="开" then%><a href="?action=ChangStat&stat=closethis&sid=<%=rs("sid")%>">封盘</a><%else%><a href="?action=ChangStat&stat=openthis&sid=<%=rs("sid")%>" class=cred>开市</a><%end if%> | <a href="?action=edit&sid=<%=rs("sid")%>">编辑</a> | <a href="?action=del&sid=<%=rs("sid")%>" onclick="javascript:{if(confirm('您确定删除 <%=rs("QiYe")%> 这个股票吗?')){return true;}return false;}" class=cred>删除</a></td>
		  </tr>
			<%page_count = page_count + 1		
			rs.movenext
		loop
		Pcount=rs.PageCount%>
		<tr>
			<td class=tablebody2 colspan=2 align=left>共有 <font color=blue><%=totalrec%></font> 支股票，每页 <font color=blue><%=rs.PageSize%></font> 支，第 <font color=red><%=currentpage%></font> 页 / 共 <font color=blue><%=Pcount%></font> 页</td>
			<td class=tablebody2 colspan=6 align=right>分页：<%call disppagenum(currentpage,Pcount,"""?page=","""")%>&nbsp;&nbsp;&nbsp;&nbsp;</td>
		</tr>
	<%end if
	rs.close
	set rs=nothing	
end sub
'-----------------------
sub ChangStat()
	on error resume next
	if request("stat")="closethis" then
		gp_conn.execute("update [GuPiao] set ZhuangTai='封' where sid="&request("sid"))
	elseif request("stat")="openthis" then
		gp_conn.execute("update [GuPiao] set ZhuangTai='开' where sid="&request("sid"))		
	end if
	response.redirect "?"
end sub
'-------------编辑股票---------------
sub Edit()
	dim Sid,ExplainSplit
	Sid=request("sid")
	if Sid="" then
		errmess="<li>参数错误，请指定相关股票"
		call endinfo(2)
		exit sub	
	end if
	set rs=gp_conn.execute("select * from [GuPiao] where sid="&Sid)
	if rs.eof and rs.bof then
		rs.close
		errmess="<li>没有找到该股票"
		call endinfo(2)
		exit sub
	end if
	ExplainSplit=split(rs("Explain"),"|")
	if ubound(ExplainSplit)<1 then
		redim preserve ExplainSplit(1)
		ExplainSplit(1)=10000
	end if%>
	<tr>
		<th colspan="2" valign=center align=middle height=25>编辑 <font color="<%=forum_body(8)%>"><%=sid%></font> 号股票<%=rs("QiYe")%></th>
	</tr>
	<form method=post  action="?action=SaveEdit">
	<tr height=25>
		<td class=tablebody2 width="40%"><b>代号</b></td>
		<td class=tablebody1 width="60%">&nbsp;<%=rs("sid")%><input type=hidden name=sid value="<%=rs("sid")%>"></td>
	</tr>
	<tr height=25>
		<td class=tablebody2><b>经营者</b></td>
		<td class=tablebody1>&nbsp;<%=rs("JingYingZhe")%></td>
	</tr>	
	<tr>
		<td class=tablebody2><b>企业</b></td>
		<td class=tablebody1>&nbsp;<input type=text name=gpname value="<%=rs("QiYe")%>"><input type=hidden name=gpoldname value="<%=rs("QiYe")%>"></td>
	</tr>
	<tr>
		<td class=tablebody2><b>总股份</b></td>
		<td class=tablebody1>&nbsp;<input type=text name=TotalStockNum value="<%=rs("ZongGuFen")%>"></td>
	</tr>
	<tr>
		<td class=tablebody2><b>初始交易量</b></td>
		<td class=tablebody1>&nbsp;<input type=text name=inigpnum value="<%=ExplainSplit(1)%>"></td>
	</tr>	
	<tr>
		<td class=tablebody2><b>剩余交易量</b></td>
		<td class=tablebody1>&nbsp;<input type=text name=gpnum value="<%=rs("JiaoYiLiang")%>"></td>
	</tr>			
	<tr>
		<td class=tablebody2><b>开盘价格</b></td>
		<td class=tablebody1>&nbsp;<input type=text name=openprice value="<%=round(rs("KaiPanJiaGe"),2)%>"></td>
	</tr>
	<tr>
		<td class=tablebody2><b>当前价格</b></td>
		<td class=tablebody1>&nbsp;<input type=text name=nowprice value="<%=round(rs("DangQianJiaGe"),2)%>"></td>
	</tr>
	<tr>
		<td class=tablebody2><b>上市日期</b></td>
		<td class=tablebody1>&nbsp;<input type=text name=gpdate value="<%=rs("RiQi")%>"></td>
	</tr>	
	<tr>
		<td class=tablebody2><b>状态</b></td>
		<td class=tablebody1><input type=radio name=state value="开" <%if rs("ZhuangTai")="开" then%> checked <%end if%>> 开<input type=radio name=state value="封" <%if rs("ZhuangTai")="封" then%> checked <%end if%>>封</td>
	</tr>
	<tr>
		<td class=tablebody2><b>公司图片</b>(150*190)<br>如果图片路径是相对路径，则当前目录是股票的根目录<br>图片也可以是网上的图片连接</td> 
		<td class=tablebody1>&nbsp;<input type=text name=LtdImg size="39" value="<%=rs("LtdImg")%>"></td> 
	</tr>	
	<tr>
		<td class=tablebody2 valign="top"><b>公司简介</b><br>上市公司介绍<br>公司简介不能超过100Byte</td>   
		<td class=tablebody1>&nbsp;<textarea cols="39" name="Explain" rows="5" wrap="PHYSICAL"><%=htmlencode(ExplainSplit(0))%></textarea></td>  
	</tr>

	<tr>
		<td class=tablebody2><font color=gray>剩余股份</font></td>
		<td class=tablebody1>&nbsp;<input type=text readonly name=remaingp value="<%=rs("ShengYuGuFen")%>"></td> 
	</tr>			
	<tr>
		<td class=tablebody2><font color=gray>今日成交量</font></td>
		<td class=tablebody1>&nbsp;<input type=text readonly name=chengjiao value="<%=rs("ChengJiaoLiang")%>"></td>
	</tr>
	<tr>
		<td class=tablebody2><font color=gray>今日买入笔数</font></td>
		<td class=tablebody1>&nbsp;<input type=text readonly name=buy value="<%=rs("MaiRuBiShu")%>"></td>
	</tr>
	<tr>
		<td class=tablebody2><font color=gray>今日卖出笔数</font></td>
		<td class=tablebody1>&nbsp;<input type=text readonly name=sale value="<%=rs("MaiChuBiShu")%>"></td>
	</tr>					
	<tr>
		<th colspan="2" align="center"><input type=submit name=submit value=执行修改></th>
	</tr>
	</form>
<%rs.close
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
		set rs=gp_conn.execute("select sid from [GuPiao] where [QiYe]='"&GpName&"'")
		if not rs.eof then
			errmess="<li>您输入的上市公司已经存在，请重新填写上市公司名称"
		end if
	end if	
	
	if errmess<>"" then
		call endinfo(2)
		exit sub	
	end if	
	if isdate(gpdate) then 
		gpdate=" ,RiQi='"&gpdate&"'"
	else
		gpdate=""
	end if
	if Explain="" then Explain="暂无"
			
	gp_conn.execute("update [GuPiao] set [QiYe]='"&GpName&"',JiaoYiLiang="&gpnum&",KaiPanJiaGe="&OpenPrice&",DangQianJiaGe="&NowPrice&",ZhuangTai='"&Request.form("state")&"',ZongGuFen="&TotalStockNum&",LtdImg='"&LtdImg&"',Explain='"&Explain&"|"&inigpnum&"' "&gpdate&" where sid="&request("sid"))
	if gpoldname<>GpName then	'改变股票名字时 request("sid")
		gp_conn.execute("update [DaHu] set QiYe='"&GpName&"' where sid="&request("sid")) 
		gp_conn.execute("update [PersonalStock] set Stockname='"&GpName&"' where Stockname='"&gpoldname&"'")
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
	set rs=gp_conn.execute("select DangQianJiaGe,QiYe from [GuPiao] where sid="&DelID)
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
	gp_conn.execute("delete from [GuPiao] where sid="&DelID)
	set rs=gp_conn.execute("select * from [DaHu] where sid="&DelID)
	do while not rs.eof
		AddMoney=rs("ChiGuShu")*NowPrice
		gp_conn.execute "update [KeHu] set ZiJin=ZiJin+"&AddMoney&",ZongZiJin=ZongZiJin-"&AddMoney&",JinRiMaiChu=JinRiMaiChu+1,ZongMaiChu=ZongMaiChu+1 where id="&rs("uid")&""
		rs.movenext
	loop
	rs.close
	gp_conn.execute("delete from [DaHu] where sid="&DelID)
	sucmess="<font color=#66CCFF>"&GPname&"</font><font color=#CCCCCC>倒闭，购买了该股的客户被迫以当前价格全部抛出</font>"
	gp_conn.execute "insert into RndEvent(content,addtime) values('"&sucmess&"','"&now()&"' )"  
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
		sql="select * from [GuPiao] where QiYe = '"&GupiaoName&"'"  
	else	
		sql="select * from [GuPiao] where QiYe like '%"&GupiaoName&"%'"         
	end if%>
	<tr align="center" > 
		<th align="center">代号</th>
		<th align="center">股票名称</th>
		<th align="center">总股份</th>
		<th align="center">昨日收盘价</th>
		<th align="center">目前成交价</th>
		<th align="center">成交量</th>
		<th align="center">状态</th>
		<th align="center">操作</th>
	</tr>
	<%set rs=gp_conn.execute(sql)         
	if rs.eof then
		response.Write "<tr><td class=tablebody1 colspan=8>没有找到任何股票</td></tr>"
	else
		do while not rs.eof%>
			<tr align="center"> 
				<td class=tablebody1><a href="z_gp_Exchange.asp?sid=<%=rs("sid")%>"> <%=rs("sid")%> </a></td>
				<td class=tablebody1><a href="?action=edit&sid=<%=rs("sid")%>"><%=rs("QiYe")%></a></td>
				<td class=tablebody1 align=right><%=formatnumber(rs("JiaoYiLiang"),0)%>&nbsp;</td>
				<td class=tablebody1 align=right><%=formatnumber(rs("KaiPanJiaGe"),2)%>&nbsp;</td>
				<td class=tablebody1 align=right><%=formatnumber(rs("DangQianJiaGe"),2)%>&nbsp;</td>
				<td class=tablebody1 align=right><%=rs("ChengJiaoLiang")%>&nbsp;</td>
				<td class=tablebody1><%if rs("ZhuangTai")="封" then%><font color=red><%end if%><%=rs("ZhuangTai")%></td>
				<td class=tablebody1><a href="?action=edit&sid=<%=rs("sid")%>">编辑</a> | <a href="?action=del&sid=<%=rs("sid")%>" onclick="javascript:{if(confirm('您确定删除 <%=rs("QiYe")%> 这个股票吗?')){return true;}return false;}">删除</a></td>
			</tr>	
  		<%rs.movenext
		loop
		rs.close
	end if%>
	<TR>
		<Td class=tablebody2 height=25 align="center" colspan="8"><A href="?">[返回列表]</A></Td>
	</TR>
<%end sub
'-----------------新股上市---------------
sub newgupiao()%>
	<tr>
		<th colspan="2" valign=center align=middle height=25>新股上市</th>
	</tr>
	<form method=post  action="?action=SaveNew">
	<tr>
		<td class=tablebody2 width="40%"><b>ID</b></td>
		<td class=tablebody1 width="60%">&nbsp;自动分配</td>
	</tr>
	<tr>
		<td class=tablebody2><b>企业(股票名)</b>字数不要超过20字</td>
		<td class=tablebody1>&nbsp;<input type=text name=gpname value=""> <font color=<%=forum_body(8)%>>*</font></td>
	</tr>
	<tr>
		<td class=tablebody2><b>初始交易量</b><br>每日交易量</td>
		<td class=tablebody1>&nbsp;<input type=text name=gpnum value="10000"> 股</td> 
	</tr>
	<tr>
		<td class=tablebody2><b>总股份</b><br>股票总数量</td> 
		<td class=tablebody1>&nbsp;<input type=text name=TotalStockNum value="10000000"> 股</td>
	</tr>				
	<tr>
		<td class=tablebody2><b>开盘价格</b></td>
		<td class=tablebody1>&nbsp;<input type=text name=openprice value="10.00"> 元</td>
	</tr>
	<tr>
		<td class=tablebody2><b>当前价格</b></td>
		<td class=tablebody1>&nbsp;<input type=text name=nowprice value="10.00"> 元</td>
	</tr>
	<tr>
		<td class=tablebody2><b>状态</b></td>
		<td class=tablebody1><input type=radio name=state value="开" checked> 开<input type=radio name=state value="封">封</td>
	</tr>
	<tr>
		<td class=tablebody2><b>公司图片</b>(150*190)<br>如果图片路径是相对路径，则当前目录是股票的根目录<br>图片也可以是网上的图片连接</td> 
		<td class=tablebody1>&nbsp;<input type=text name=LtdImg size="39" value="Images/LtdImg/1.jpg"></td> 
	</tr>	
	<tr>
		<td class=tablebody2 valign="top"><b>公司简介</b><br>上市公司介绍<br>公司简介不能超过100Byte</td>   
		<td class=tablebody1>&nbsp;<textarea cols="39" name="Explain" rows="5" wrap="PHYSICAL">这个家伙很懒，什么都没有留下！</textarea></td>  
	</tr>	
	<tr>
		<td class=tablebody2><b>成交量</b></td>
		<td class=tablebody1>&nbsp;0</td>
	</tr>	
	<tr>
		<td class=tablebody2><b>涨跌</b></td>
		<td class=tablebody1>&nbsp;0</td>
	</tr>	
	<tr>
		<td class=tablebody2><b>买入笔数</b></td>
		<td class=tablebody1>&nbsp;0</td>
	</tr>
	<tr>
		<td class=tablebody2><b>卖出笔数</b></td>
		<td class=tablebody1>&nbsp;0</td>
	</tr>					
	<tr><th colspan="2" align="center"><input type=submit name=submit value=增加><input type=reset name=submit value=重填></th>
	</form>
<%end sub
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
		errmess=errmess+"<li>每日交易量必须输入数字"
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
	set rs=gp_conn.execute("select sid from [GuPiao] where QiYe='"&GpName&"'")
	if not rs.eof then
		errmess="<li>您输入的股票名称已经上市了，请重新填写股票名称"
		endinfo(2)
		exit sub
	else
		gp_conn.execute("insert into [GuPiao](QiYe,JiaoYiLiang,KaiPanJiaGe,DangQianJiaGe,ZhuangTai,RiQi,ZongGuFen,LtdImg,Explain,TongJi) values('"&GpName&"',"&gpnum&","&OpenPrice&","&NowPrice&",'"&Request.form("state")&"',date(),"&TotalStockNum&",'"&LtdImg&"','"&Explain&"|"&gpnum&"','0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0')")
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
	set rs=gp_conn.execute("select sid from [GuPiao] order by sid")
	Cid=1
	do while not rs.eof
		gp_conn.execute("update [GuPiao] set Cid="&Cid&" where sid="&rs(0))
		Cid=Cid+1	
		rs.movenext
	loop
end sub
%>