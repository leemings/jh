<!--#include file=conn.asp-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_visual_const.asp" -->
<%
dim CurVisual,CurVisualSplit,TempSplit,ItemName,rsvisual,ItemCount
dim j,flag,SucFlag,buyitems,buycount
dim myvisualsplit

founderr=false
myvisualsplit=split(v_myvisual,"|")

if not founduser or membername=""  then
	errmsg=errmsg+"<br>"+"<li>您还没有<a href=login.asp>登录论坛</a>，不能使用个人形象设计功能。如果您还没有<a href=reg.asp>注册</a>，请先<a href=reg.asp>注册</a>！"
	founderr=true
end if
CurVisual=request.cookies("myshow_"&userid)
if isnull(CurVisual) then CurVisual=""
CurVisualSplit=split(CurVisual,"|")
if CurVisual="" or ubound(CurVisualSplit)<>24 then
	errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
	founderr=true
end if
stats="保存形象"
if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	set rsvisual=server.createobject("ADODB.Recordset")
	SucFlag=true
	ItemCount=0
	ItemName=""
	BuyItems=""
	BuyCount=0
	for i=0 to ubound(CurVisualSplit)
		if myVisualSplit(i)<>"" then
			TempSplit=split(myVisualSplit(i),".")(0)
		else
			TempSplit=""
		end if
		if CurVisualSplit(i)<>TempSplit then
			if CurVisualSplit(i)="" then
				if cint(TempSplit)>20 then
					call DressDown()
				end if
			elseif TempSplit="" then
				if cint(CurVisualSplit(i))>20 then
					sql="select top 1 fixdate,period from visual_myitems where username='"&membername&"' and itemid="&CurVisualSplit(i)
					rsvisual.open sql,conn,1,1
					if not (rsvisual.eof or rsvisual.bof) then
						call DressUp()
					else
						SucFlag=false
						call DressBuy()
					end if
					rsvisual.close
				end if
			else
				if cint(CurVisualSplit(i))>20 then
					if cint(TempSplit)>20 then
						call DressDown()
					end if
					sql="select top 1 fixdate,period from visual_myitems where username='"&membername&"' and itemid="&CurVisualSplit(i)
					rsvisual.open sql,conn,1,1
					if not (rsvisual.eof or rsvisual.bof) then
						call DressUp()
					else
						SucFlag=false
						call DressBuy()
					end if
					rsvisual.close
				else
					if cint(TempSplit)>20 then
						call DressDown()
					end if
				end if
			end if
		end if
	next
	set rsvisual=nothing
	if SucFlag and ItemName<>"" then
		curvisual=""
		set rs=server.createobject("adodb.recordset")
		set rsvisual=server.createobject("adodb.recordset")
		for i=0 to 24
			if i>0 then curvisual=curvisual&"|"
			if curvisualsplit(i)<>"" then
				rs.open "select * from visual_items where id="&curvisualsplit(i),conn,1,1
				if cint(curvisualsplit(i))>20 then
					rsvisual.open "select * from visual_myitems where itemid="&curvisualsplit(i)&" and username='"&membername&"'",conn,1,3
					if rsvisual("type")=5 then
						if rs("type") mod 100>=5 then
							rsvisual("type")=4
						else
							rsvisual("type")=rs("type") mod 100
						end if
						rsvisual.update
					end if
					rsvisual.close
				end if
				curvisual=curvisual&curvisualsplit(i)&"."&rs("content")
				rs.close
			end if
		next
		set rs=nothing
		set rsvisual=nothing
	end if
	if SucFlag and ItemName<>"" then
		conn.execute("update [user] set visual='"&curvisual&"' where username='"&membername&"'")
		response.cookies("myshow_"&userid)=""
		sucmsg="<li>您成功地："
		sucmsg=sucmsg+ItemName
		call nav()
		call head_var(2,0,"","")
		call dvbbs_suc()
	elseif not SucFlag and BuyItems<>"" then
		errmsg=errmsg+"<br>"+"<li>如下商品尚未购买不能换上："
		errmsg=errmsg+BuyItems
		founderr=true
		call nav()
		call head_var(2,0,"","")
		response.write "<script language=javascript>openScript('z_visual_cartbag.asp',500,440);</script>"
		call dvbbs_error()
	else
		errmsg=errmsg+"<br>"+"<li>您没有更换任何商品！"
		founderr=true
		call nav()
		call head_var(2,0,"","")
		call dvbbs_error()
	end if
end if
call footer()

sub DressDown()
	dim flag,j
	
	flag=false
	for j=0 to i-1
		if MyVisualSplit(i)=MyVisualSplit(j) then
			flag=true
			exit for
		end if
	next
	if not flag then
		ItemCount=ItemCount+1
		ItemName=ItemName&"<br>&nbsp;&nbsp;&nbsp;"&ItemCount&"、换下了<font color=red>"&conn.execute("select top 1 name from visual_items where id="&TempSplit)(0)&"</font>"
	end if
end sub

sub DressUp()
	dim flag,j
	
	flag=false
	for j=0 to i-1
		if CurVisualSplit(j)<>"" then
			if CurVisualSplit(i)=CurVisualSplit(j) then
				flag=true
				exit for
			end if
		end if
	next
	if not flag then
		ItemCount=ItemCount+1
		ItemName=ItemName&"<br>&nbsp;&nbsp;&nbsp;"&ItemCount&"、换上了<font color=blue>"&conn.execute("select top 1 name from visual_items where id="&CurVisualSplit(i))(0)&"</font>"
	end if
end sub

sub DressBuy()
	dim flag,j
	dim CartBag,rsitems
	
	flag=false
	for j=0 to i-1
		if CurVisualSplit(j)<>"" then
			if CurVisualSplit(i)=CurVisualSplit(j) then
				flag=true
				exit for
			end if
		end if
	next
	if not flag then
		BuyCount=BuyCount+1
		BuyItems=BuyItems&"<br>&nbsp;&nbsp;&nbsp;"&BuyCount&"、"&conn.execute("select top 1 name from visual_items where id="&CurVisualSplit(i))(0)
		set rsitems=server.createobject("ADODB.recordset")
		sql="select * from visual_items where id="&CurVisualSplit(i)
		rsitems.open sql,conn,1,1
		flag=true
		if rsitems("vip") and v_myvip<>1 then
			BuyItems=BuyItems&"&nbsp;★您不是VIP，不能购买VIP专用商品！"
			flag=false
		end if
		if rsitems("quantity")<=0 then
			BuyItems=BuyItems&"&nbsp;★该商品库存不足，您无法购买！"
			flag=false
		end if
		if rsitems("flag")=0 then
			BuyItems=BuyItems&"&nbsp;★保留商品，您无法购买！"
			flag=false
		end if
		if rsitems("flag")=1 and membername="" then
			BuyItems=BuyItems&"&nbsp;★只有会员才能购买此商品，您无法购买！"
			flag=false
		end if
		if rsitems("flag")=2 and not master and not boardmaster and not superboardmaster then
			BuyItems=BuyItems&"&nbsp;★只有版主以上人员才能购买此商品，您无法购买！"
			flag=false
		end if
		if rsitems("flag")=3 and not master and not superboardmaster then
			BuyItems=BuyItems&"&nbsp;★只有超级版主和管理员才能购买此商品，您无法购买！"
			flag=false
		end if
		if rsitems("flag")=4 and not master then
			BuyItems=BuyItems&"&nbsp;★只有管理员才能购买此商品，您无法购买！"
			flag=false
		end if
		if flag then
			CartBag=request.cookies("mybuy_"&userid)
			if isnull(CartBag) then CartBag=""
			if instr(1,"|"&CartBag&"|","|"&CurVisualSplit(i)&"|")<=0 then
				if ubound(split(CartBag,"|"))<CartLength-1 then
					if CartBag<>"" then CartBag=CartBag&"|"
					CartBag=CartBag&CurVisualSplit(i)
					response.cookies("mybuy_"&userid)=CartBag
					BuyItems=BuyItems&"&nbsp;★已将此商品放入您的购物袋！"
				else
					BuyItems=BuyItems&"&nbsp;★您的购物袋已满，无法将此商品放入！"
				end if
			else
				BuyItems=BuyItems&"&nbsp;★购物袋中已有此商品，无须再次放入！"
			end if
		end if
	end if
end sub
%>