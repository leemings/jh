<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<%
dim CurVisual,CurVisualSplit,TempSplit,TotalAmount,CurPeriod,CurPeriodSplit(24),ItemName,rsvisual,ItemCount,ItemPrice
dim j,flag,flag2,succflag,buyitems,buyitemsplit
dim myLayerStr,myvisualsplit,myperiodsplit

myvisualsplit=split(myvisual,"|")
myLayerStr=GetLayerStr(myvisualsplit)
myperiodsplit=split(myperiod,"|")

if not founduser or membername=""  then
	errmsg=errmsg+"<br>"+"<li>您还没有<a href=login.asp>登录论坛</a>，不能使用个人形象设计功能。如果您还没有<a href=reg.asp>注册</a>，请先<a href=reg.asp>注册</a>！"
	founderr=true
end if
CurVisual=request.cookies("myshow_"&userid)
if ubound(myvisualsplit)<>24 or CurVisual="" then
	errmsg=errmsg+"<br>"+"<li>您没有购买任何商品！"
	founderr=true
end if
if isnull(Request.ServerVariables("HTTP_REFERER")) or instr(1,lcase(Request.ServerVariables("HTTP_REFERER")),"/z_visual.asp")<=0 then
	errmsg=errmsg+"<br>"+"<li>请从个人形象设计页面提交数据！"
	founderr=true
end if
CurVisualSplit=split(CurVisual,"|")
if CurVisual<>"" and ubound(CurVisualSplit)<>24 then
	errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
	founderr=true
end if

if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	totalamount=0
	CurPeriod=""
	ItemName=""
	ItemCount=0
	set rsvisual=server.createobject("ADODB.Recordset")
	flag2=false
	succflag=true
	buyitems=""
	for i=0 to ubound(CurVisualSplit)
		CurPeriodSplit(i)=myperiodsplit(i)
		if myVisualSplit(i)<>"" then
			TempSplit=split(myVisualSplit(i),".")(0)
		else
			TempSplit=""
		end if
		if CurVisualSplit(i)<>TempSplit then
			if CurVisualSplit(i)<>"" then
				if cint(CurVisualSplit(i))>20 then
					flag=false
					for j=0 to i-1
						if CurVisualSplit(j)<>"" then
							if CurVisualSplit(i)=CurVisualSplit(j) then
								flag=true
								exit for
							end if
						end if
					next
					if flag then
						CurPeriodSplit(i)=CurPeriodSplit(j)
					else
						sql="select * from visual_items where id="&CurVisualSplit(i)
						rsvisual.open sql,conn,1,1
						if not (rsvisual.eof or rsvisual.bof) then
							CurPeriodSplit(i)=right("00"&(year(now) mod 100),2)&right("00"&month(now),2)&right("00"&day(now),2)&","&rsvisual("period")
							if isnull(myvip) or myvip<>1 then
								ItemPrice=rsvisual("price1")
							else
								ItemPrice=rsvisual("price2")
							end if
							totalamount=totalamount+ItemPrice
							ItemCount=ItemCount+1
							ItemName=ItemName&"<br>&nbsp;&nbsp;&nbsp;"&ItemCount&"、"&rsvisual("name")&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&ItemPrice&" 元"
							if buyitems<>"" then buyitems=buyitems&"|"
							buyitems=buyitems&CurVisualSplit(i)
							if rsvisual("vip") and (isnull(myvip) or myvip<>1) then
								ItemName=ItemName&"&nbsp;★您不是VIP，不能购买VIP专用商品"
								succflag=false
							end if
							if rsvisual("quantity")<=0 then
								ItemName=ItemName&"&nbsp;★该商品库存不足，您无法购买"
								succflag=false
							end if
							if rsvisual("flag")=0 then
								ItemName=ItemName&"&nbsp;★保留商品，您无法购买"
								succflag=false
							end if
							if rsvisual("flag")=1 and membername="" then
								ItemName=ItemName&"&nbsp;★只有会员才能购买此商品，您无法购买"
								succflag=false
							end if
							if rsvisual("flag")=2 and not master and not boardmaster and not superboardmaster then
								ItemName=ItemName&"&nbsp;★只有版主以上人员才能购买此商品，您无法购买"
								succflag=false
							end if
							if rsvisual("flag")=3 and not master and not superboardmaster then
								ItemName=ItemName&"&nbsp;★只有超级版主和管理员才能购买此商品，您无法购买"
								succflag=false
							end if
							if rsvisual("flag")=4 and not master then
								ItemName=ItemName&"&nbsp;★只有管理员才能购买此商品，您无法购买"
								succflag=false
							end if
						end if
						rsvisual.close
					end if
				end if
			end if
			flag2=true
		end if
		if i>0 then CurPeriod=CurPeriod&"|"
		CurPeriod=CurPeriod&CurPeriodSplit(i)
	next
	set rsvisual=nothing
	curvisual=""
	for i=0 to 24
		if i>0 then curvisual=curvisual&"|"
		if curvisualsplit(i)<>"" then
			curvisual=curvisual&curvisualsplit(i)&"."&(i+1)
		end if
	next
	if totalamount>mymoney then	
		errmsg=errmsg+"<br>"+"<li>下列商品共计需要"&totalamount&"元，您只有"&mymoney&"元，您的现金不够买这些商品："
		errmsg=errmsg+ItemName
		founderr=true
		call nav()
		call head_var(2,0,"","")
		call dvbbs_error()
	elseif flag2 and succflag and ItemName<>"" then
		conn.execute("update [user] set userwealth=userwealth-"&totalamount&",visual='"&curvisual&"',Period='"&CurPeriod&"' where username='"&membername&"'")
		buyitemsplit=split(buyitems,"|")
		for i=0 to ubound(buyitemsplit)
			conn.execute("update visual_items set quantity=quantity-1,totalsales=totalsales+1 where id="&buyitemsplit(i))
		next
		response.cookies("myshow_"&userid)=""
		sucmsg="<li>您总计花费"&totalamount&"元成功地购买了以下商品："
		sucmsg=sucmsg+ItemName
		call nav()
		call head_var(2,0,"","")
		call dvbbs_suc()
	elseif flag2 and ItemName<>"" then
		errmsg="<li>由于以下原因，您的购买没有成功："
		errmsg=errmsg+ItemName
		call nav()
		call head_var(2,0,"","")
		call dvbbs_error()
	elseif ItemName="" then
		conn.execute("update [user] set visual='"&curvisual&"',Period='"&CurPeriod&"' where username='"&membername&"'")
		response.cookies("myshow_"&userid)=""
		sucmsg="<li>您成功地重新设计了自己的形象！"
		call nav()
		call head_var(2,0,"","")
		call dvbbs_suc()
	else
		errmsg=errmsg+"<br>"+"<li>您没有购买任何商品！"
		founderr=true
		call nav()
		call head_var(2,0,"","")
		call dvbbs_error()
	end if
end if
call footer()
%>