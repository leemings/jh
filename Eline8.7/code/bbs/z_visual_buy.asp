<!--#include file=conn.asp-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_visual_const.asp" -->
<%
dim CurItems,ItemName,rsvisual
dim SucFlag,CurItem,CartBag,SendMsg,CurType,SalePrice

founderr=false
if not founduser or membername=""  then
	sendmsg=sendmsg+"<br>"+"<li>您还没有<a href=login.asp>登录论坛</a>，不能使用个人形象设计功能。如果您还没有<a href=reg.asp>注册</a>，请先<a href=reg.asp>注册</a>！"
	founderr=true
end if
CurItem=CheckStr(request("itemid"))
if isnull(curitem) then
	sendmsg=sendmsg+"<br>"+"<li>没有指定商品编号！"
	founderr=true
else
	if not isnumeric(curitem) then
		sendmsg=sendmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
	end if
end if

if founderr then
	call send_head()
	call send_msg()
else
	set rsvisual=server.createobject("ADODB.Recordset")
	sql="select top 1 itemid from visual_myitems where username='"&membername&"' and itemid="&CurItem
	rsvisual.open sql,conn,1,1
	if not (rsvisual.eof or rsvisual.bof) then
		sendmsg=sendmsg+"<br>"+"<li>已经拥有该商品，无须再次购买！"
		founderr=true
	else
		if request("fromuser")<>"" then
			rsvisual.close
			sql="select m.username,m.saleprice,v.name,v.type from visual_myitems m inner join visual_items v on m.itemid=v.id where m.itemid="&CurItem&" and m.username='"&checkstr(trim(request("fromuser")))&"'"
			rsvisual.open sql,conn,1,1
			ItemName=rsvisual("name")
			CurType=rsvisual("type") mod 100
			if CurType>=5 then CurType=4
			saleprice=rsvisual("salePrice")
			if saleprice>mymoney then
				sendmsg=sendmsg+"<br>"+"<li>购买此商品需要现金 "&saleprice&" 元，您的现金("&mymoney&" 元)不足，无法购买！"
				founderr=true
			end if
		else
			rsvisual.close
			sql="select * from visual_items where id="&CurItem
			rsvisual.open sql,conn,1,1
			ItemName=rsvisual("name")
			if rsvisual("vip") and v_myvip<>1 then
				sendmsg=sendmsg+"<br>"+"<li>您不是VIP，不能购买VIP专用商品！"
				founderr=true
			end if
			if rsvisual("quantity")<=0 then
				sendmsg=sendmsg+"<br>"+"<li>该商品库存不足，您无法购买！"
				founderr=true
			end if
			if rsvisual("flag")=0 then
				sendmsg=sendmsg+"<br>"+"<li>保留商品，您无法购买！"
				founderr=true
			end if
			if rsvisual("flag")=1 and membername="" then
				sendmsg=sendmsg+"<br>"+"<li>只有会员才能购买此商品，您无法购买！"
				founderr=true
			end if
			if rsvisual("flag")=2 and not master and not boardmaster and not superboardmaster then
				sendmsg=sendmsg+"<br>"+"<li>只有版主以上人员才能购买此商品，您无法购买！"
				founderr=true
			end if
			if rsvisual("flag")=3 and not master and not superboardmaster then
				sendmsg=sendmsg+"<br>"+"<li>只有超级版主和管理员才能购买此商品，您无法购买！"
				founderr=true
			end if
			if rsvisual("flag")=4 and not master then
				sendmsg=sendmsg+"<br>"+"<li>只有管理员才能购买此商品，您无法购买！"
				founderr=true
			end if
		end if
		if founderr then
			call send_head()
			call send_msg()
		else	
			if request("fromuser")<>"" then
				conn.execute("update visual_myitems set type="&curtype&",price=saleprice,saleprice=0,fromuser=username,username='"&membername&"' where itemid="&curitem&" and type=5 and username='"&checkstr(trim(request("fromuser")))&"'")
				conn.execute("update visual_events set username='"&membername&"' where itemid="&curitem&" and type=2 and fromuser='"&checkstr(trim(request("fromuser")))&"' and (username='' or isnull(username))")
				conn.execute("insert into visual_events (ItemID,UserName,FromUser,DateAndTime,Price,Type) values ("&CurItem&",'"&membername&"','"&checkstr(trim(request("fromuser")))&"','"&year(now())&"-"&month(now())&"-"&day(now())&"',"&saleprice&",0)")
				conn.execute("update [user] set userwealth=userwealth-"&saleprice&" where username='"&membername&"'")
				conn.execute("update [user] set userwealth=userwealth+"&saleprice&" where username='"&checkstr(trim(request("fromuser")))&"'")
				set rs=server.createobject("adodb.recordset")
				sql="select * from [message]"      
				rs.open sql,conn,1,3      
				rs.addnew      
				rs("sender")=forum_info(0)
				rs("incept")=trim(request("fromuser"))
				rs("title")="您出售的商品有人购买了！"
				rs("content")="您在二手市场出售的商品《[color=blue]"&itemName&"[/color]》已经被"&membername&"购买了！"
				rs("flag")=0 
				rs("issend")=1
				rs.update
				rs.close
				set rs=nothing
				sendmsg="<li>已经成功地购买了二手商品 <font color=blue>"&ItemName&"</font>！"
				call send_head()
				response.write "<script>"
				response.write "window.opener.document.execCommand('Refresh');"
				response.write "</script>"
				call send_msg()
			else
				CartBag=request.cookies("mybuy_"&userid)
				if isnull(CartBag) then CartBag=""
				if instr(1,"|"&CartBag&"|","|"&CurItem&"|")<=0 then
					if ubound(split(CartBag,"|"))<CartLength-1 then
						if CartBag<>"" then CartBag=CartBag&"|"
						CartBag=CartBag&CurItem
						response.cookies("mybuy_"&userid)=CartBag
						sendmsg="<li>已将商品 <font color=blue>"&ItemName&"</font> 放到您的购物袋中！"
					else
						sendmsg=sendmsg+"<br>"+"<li>您的购物袋已满，无法将此商品放入！"
						founderr=true
					end if
				else
					sendmsg=sendmsg+"<br>"+"<li>购物袋中已有此商品，无须再次购买！"
					founderr=true
				end if
				call send_head()
				call send_msg()
			end if
		end if
	end if
	rsvisual.close
	set rsvisual=nothing
end if
call footer()

sub send_head()%>
	<html><head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title><%=Forum_info(0)%>--购买形象商品</title>
	<!--#include file=inc/forum_css.asp-->
	</head>
	<body <%=Forum_body(11)%>">
<%end sub

sub send_msg()%>
	<br>
	<table cellpadding=3 cellspacing=1 align=center bgcolor=<%=Forum_Body(27)%> width=97%>
		<tr align=center>
			<th width="100%">购买形象商品信息</th>
		</tr>
		<tr>
			<td width="100%" class=tablebody1><b>操作结果：</b><br><%
				if founderr then
					response.write "<font color="&forum_body(8)&">"
					response.write sendmsg
					response.write "</font>"
				else
					response.write sendmsg
				end if
			%></td>
		</tr>
		<tr align=center>
			<td width="100%" class=tablebody2><a href="javascript:window.close()">[关闭窗口]</a></td>
		</tr>
	</table>
<%end sub%>