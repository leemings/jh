<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_visual_const.asp"-->
<%
dim CartBag,SendSplit,SumAmount,SendMsg,rs2,sendfee
dim itemsplit,itemcount,rsvisual,curtype,CurPrice,CurDate
founderr=false
CartBag=request.cookies("mybuy_"&userid)
if isnull(CartBag) then CartBag=""
if CartBag="" then
	itemcount=0
else
	itemsplit=split(CartBag,"|")
	itemcount=ubound(itemsplit)+1
end if

if request.form("cmdclear")="清空" then
	response.cookies("mybuy_"&userid)=""
	response.write "<script>"
	response.write "window.close();"
	response.write "</script>"
	response.end
end if
if request.form("cmddelete")="删除" then
	if itemcount>0 then
		if not isnull(request("seqno")) then
			if isnumeric(request("seqno")) then
				if cint(request("seqno"))<itemcount then
					CartBag=""
					for i=0 to itemcount-1
						if i<>cint(request("seqno")) then
							if CartBag<>"" then CartBag=CartBag&"|"
							CartBag=CartBag&itemsplit(i)
						end if
					next
					response.cookies("mybuy_"&userid)=CartBag
					response.redirect Request.ServerVariables("HTTP_REFERER")
				end if
			end if
		end if
	end if
end if
if request.form("cmdclear")="结算" then
	set rsvisual=server.createobject("ADODB.Recordset")
	for i=0 to itemcount-1
		SendSplit=split(itemsplit(i),"$")
		sql="select top 1 * from visual_items where id="&SendSplit(0)
		rsvisual.open sql,conn,1,1
		sendfee=int(rsvisual("price1")*conn.execute("select top 1 send from visual_config")(0)/100+0.99)
		if v_myvip<>1 then
  		SumAmount=SumAmount+rsvisual("price1")
  	else
  		SumAmount=SumAmount+rsvisual("price2")
  	end if
  	if ubound(SendSplit)=2 then
  		SumAmount=SumAmount+SendFee
  	end if
		rsvisual.close
	next
	if mymoney<SumAmount then
		SendMsg=SendMSg+"<br>"+"<li>购买的商品共计 "&SumAmount&" 元，您的现金 "&mymoney&" 元不够买这些商品。"
		founderr=true
	else
		for i=0 to itemcount-1
			SendSplit=split(itemsplit(i),"$")
			sql="select top 1 * from visual_items where id="&SendSplit(0)
			rsvisual.open sql,conn,1,1
			sendfee=int(rsvisual("price1")*conn.execute("select top 1 send from visual_config")(0)/100+0.99)
			if v_myvip<>1 then
				CurPrice=rsvisual("price1")
	  	else
				CurPrice=rsvisual("price2")
	  	end if
			curtype=rsvisual("type") mod 100
			if curtype>=5 then
				curtype=4
			end if
			CurDate=year(now())&"-"&month(now())&"-"&day(now())
			conn.execute("update visual_items set quantity=quantity-1,totalsales=totalsales+1 where id="&SendSplit(0))
			if ubound(SendSplit)<2 then
				conn.execute("insert into visual_myitems (ItemID,UserName,FromUser,[Type],AddDate,FixDate,Period,Price) values ("&SendSplit(0)&",'"&membername&"','"&membername&"',"&curtype&",'"&curdate&"','"&curdate&"',"&rsvisual("period")&","&curprice&")")
				conn.execute("insert into visual_events (ItemID,UserName,FromUser,DateAndTime,Price,Type) values ("&sendsplit(0)&",'"&membername&"','#形象商店#','"&curdate&"',"&curprice&",0)")
			else
				conn.execute("insert into visual_myitems (ItemID,UserName,FromUser,[Type],AddDate,FixDate,Period,Price) values ("&SendSplit(0)&",'"&SendSplit(1)&"','"&membername&"',"&curtype&",'"&curdate&"','"&curdate&"',"&rsvisual("period")&","&curprice&")")
				conn.execute("insert into visual_events (ItemID,UserName,FromUser,DateAndTime,Price,Type) values ("&sendsplit(0)&",'"&SendSplit(1)&"','"&membername&"','"&curdate&"',"&curprice&",1)")
				set rs2=server.createobject("adodb.recordset")
				sql="select * from [message]"      
				rs2.open sql,conn,1,3      
				rs2.addnew      
				rs2("sender")=membername      
				rs2("incept")=SendSplit(1)
				rs2("title")="送您一件礼物，《"&rsvisual("name")&"》！"
				rs2("content")=SendSplit(2)&chr(10)&chr(10)&"商品已经放到您的储物柜中，请笑纳！"&chr(10)&"请到以下地址查看储物柜。"&chr(10)&"[align=center][URL=z_visual.asp?shopid=100][color=blue][b][u]查看储物柜[/u][/b][/color][/URL][/align]"
				rs2("flag")=0 
				rs2("issend")=1
				rs2.update
				rs2.close
				set rs2=nothing
			end if
	  	if ubound(SendSplit)=2 then
	  		CurPrice=CurPrice+SendFee
	  	end if
  		conn.execute("update [user] set userwealth=userwealth-"&CurPrice&" where username='"&membername&"'")
			rsvisual.close
		next
		response.cookies("mybuy_"&userid)=""
		SendMsg=SendMSg+"<br>"+"<li>您购买的商品已经放入您的储物柜中，赠送的商品已经发送到指定人员的储物柜。"
		call cartbag_head()
		response.write "<script>"
		response.write "window.opener.document.URL='z_visual_save.asp';"
		'response.write "window.opener.document.execCommand('Refresh');"
		response.write "</script>"
		call send_msg()
		call activeonline()
		response.end
	end if
	set rsvisual=nothing
end if
if not founduser then
	SendMsg=SendMSg+"<br>"+"<li>您没有<a href=login.asp target=_blank>登录</a>。"
	founderr=true
end if
stats="查看我的购物袋"
if founderr then
	call cartbag_head()
	call send_msg()
else
	call cartbag_head()
	call main()
end if
call activeonline()
call footer()

sub main()
	SumAmount=0%>
	<table width=95% cellpadding=3 cellspacing=1 border=0 align=center bgcolor=<%=Forum_Body(27)%> width=97%>
		<tbody>
		<tr height=25>
			<th colspan=5><font color=<%=forum_body(8)%>><%=membername%> ( <%=mymoney%> 元 )</font>：形象商品--购物袋</th>
		</tr>
		<tr height=20 valign=middle>
			<td align=center class=tablebody2><b>商品名称</b></td>
			<td align=center class=tablebody2><b>享受单价</b></td>
			<td align=center class=tablebody2><b>受赠方</b></td>
			<td align=center class=tablebody2><b>操作</b></td>
			<td id=itempic align=center class=tablebody2 rowspan="<%
				if itemcount=0 then
					response.write 2
				else
					response.write itemcount+1
				end if
			%>" baseName="images/img_visual/blank.gif" width="84" height="84" border="0" itemgender="" itemno="" layerno="" localpic="<%=LocalPic%>" ImagePath="<%=PicPath%>" style="behavior:url(inc/z_show2.htc)"></td>
		</tr>
		<form action=? method=post name=form1>
		<input type=hidden name=seqno value="">
		<%if itemcount=0 then%>
			<tr height=70 valign=middle>
				<td colspan=4 align=center class=tablebody1>没有购买任何商品！</td>
			</tr>
		<%else
			set rsvisual=server.createobject("ADODB.Recordset")
			for i=0 to itemcount-1
				SendSplit=split(itemsplit(i),"$")
				sql="select top 1 * from visual_items where id="&SendSplit(0)
				rsvisual.open sql,conn,1,1
				sendfee=int(rsvisual("price1")*conn.execute("select top 1 send from visual_config")(0)/100+0.99)%>
				<tr height=20 valign=middle itemno="<%=SendSplit(0)%>">
					<td align=left class=tablebody1 onmouseover=moveoverobject(this.parentElement,'TableBody2') onmouseout=moveoutobject(this.parentElement,'TableBody1')>&nbsp;<%
        		response.write rsvisual("name")
          	response.write "("
          	if rsvisual("sex")=1 then
          		response.write "男"
          	elseif rsvisual("sex")=0 then
          		response.write "女"
          	else
          		response.write "不限"
	          end if
	          response.write ")"
					%></td>
					<td align=right class=tablebody1 onmouseover=moveoverobject(this.parentElement,'TableBody2') onmouseout=moveoutobject(this.parentElement,'TableBody1')><%
						if v_myvip<>1 then
          		if rsvisual("price1")<=0 then
          			response.write "<font color="&forum_body(8)&">免费</font>"
          		else
          			response.write rsvisual("price1")&" 元"
          		end if
          		SumAmount=SumAmount+rsvisual("price1")
          	else
          		if rsvisual("price2")<=0 then
          			response.write "<font color="&forum_body(8)&">免费</font>"
          		else
          			response.write rsvisual("price2")&" 元"
          		end if
          		SumAmount=SumAmount+rsvisual("price2")
          	end if
				  	if ubound(SendSplit)=2 then
				  		response.write " <font color="&forum_body(8)&"> + "&SendFee&" 元</font>"
				  		SumAmount=SumAmount+SendFee
				  	end if
					%>&nbsp;</td>
					<td align=center class=tablebody1 onmouseover=moveoverobject(this.parentElement,'TableBody2') onmouseout=moveoutobject(this.parentElement,'TableBody1')><%
						if ubound(SendSplit)=2 then
							response.write "<a href=z_visual_send.asp?itemid="&SendSplit(0)&"&username="&SendSplit(1)&"><font color=blue>"&SendSplit(1)&"</font></a>"
						else
							response.write "自用"
						end if
					%></td>
					<td align=center class=tablebody1 onmouseover=moveoverobject(this.parentElement,'TableBody2') onmouseout=moveoutobject(this.parentElement,'TableBody1')><input type="submit" name="cmddelete" value="删除" onclick="if(confirm('您确定删除这个商品吗?')){this.document.form1.seqno.value='<%=i%>';this.document.form1.submit();return true;}return false;"></td>
				</tr>
				<%rsvisual.close
			next
			set rsvisual=nothing
		end if%>
		<tr height=20 valign=middle>
			<td class=tablebody2 align=center><b>合计</b></td>
			<td class=tablebody2 align=right><b><%=SumAmount%> 元</b>&nbsp;</td>
			<td class=tablebody2 align=center><%
				if itemcount=0 then
					%>&nbsp;<%
				else
					%><input type="submit" name="cmdclear" value="清空" onclick="if(confirm('您确定清空购物袋吗?')){this.document.form1.submit();return true;}return false;"><%
				end if
			%></td>
			<td class=tablebody2 align=center><%
				if itemcount=0 then
					%>&nbsp;<%
				else
					%><input type="submit" name="cmdclear" value="结算" onclick="if(confirm('您确定结算这些商品吗?')){this.document.form1.submit();return true;}return false;"><%
				end if
			%></td>
			<td class=tablebody2 align=center><input type="button" name="cmdclose" value="关闭" onclick="window.close();"></td>
		</tr>
		</form>
		</tbody>
	</table>
	<script>
		function moveoverobject(object,classname) {
			for(var i=0;i<object.children.length;i++) 
				object.children[i].className=classname;
			if(itempic.readyState=="complete") itempic.ShowThisIMG(object.itemno);
		}
		function moveoutobject(object,classname) {
			for(var i=0;i<object.children.length;i++) 
				object.children[i].className=classname;
			if(itempic.readyState=="complete") itempic.ShowThisIMG('');
		}
	</script>
<%end sub

sub CartBag_Head()%>
	<html>
	<head>
	<title><%=Forum_info(0)%>--我的购物袋</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<!--#include file="inc/Forum_css.asp"-->
	</head>
	<body <%=Forum_body(11)%>>
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
			<td width="100%" class=tablebody2><a href="javascript:window.close()">[关闭窗口]</a><%if founderr then%>&nbsp;&nbsp;<a href="javascript:history.go(-1)">[返回上页]</a><%end if%></td>
		</tr>
	</table>
<%end sub%>