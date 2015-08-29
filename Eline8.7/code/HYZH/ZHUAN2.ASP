<%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
username=Session("sjjh_name")
hymonth=1
hygrade=2
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	conn.open Application("sjjh_usermdb")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	sqlstr="SELECT * FROM 用户 where 姓名='"&username&"'"
	rs.Open sqlstr,conn,1,2
	if rs.EOF or rs.BOF then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('无此用户！请查看是否正确！');history.back();</script>"
		Response.End
	end if
        if rs("会员金卡")<150 then
		Response.Write "<script language=javascript>alert('抱歉你的条件不够，转换需要多加100的手续费！共需金卡150元，金币150个，银子2亿，充值卡2张');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
		if rs("金币")<150 then
		Response.Write "<script language=javascript>alert('抱歉你的条件不够，转换需要多加20的手续费！金币150个不够');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
	    if rs("银两")<200000000 then
		Response.Write "<script language=javascript>alert('抱歉你的条件不够，转换需要多加20的手续费！银子2亿不够');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
	    if rs("充值卡")<2 then
		Response.Write "<script language=javascript>alert('抱歉你的条件不够，充值卡2张不够');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
    end if
	if rs("会员等级")<>0 or rs("grade")>6 then
		Response.Write "<script language=javascript>alert('抱歉你已经是["&rs("会员等级")&"级]江湖会员！');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
	hydate=date()
	n=Year(hydate)
	y=Month(hydate)
	r=Day(hydate)
	y=y+hymonth
	if y>12 then 
		y=y-12
		n=n+1
	end if
	jy=500000
	
	sj=n & "-" & y & "-" & r
	
	rs("会员等级")=hygrade
	rs("会员日期")=cdate(sj)
	rs("allvalue")=rs("allvalue")+jy
	rs.Update
	conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','操作','[兑换"&hygrade&"级会员]日期："&sj&"')"
	conn.execute "update 用户 set 会员金卡=会员金卡-150,金币=金币-150,银两=银两-200000000,充值卡=充值卡-2 where 姓名='"&username&"'"
        set rs=nothing
	conn.Close
	set conn=nothing
	Response.Write "<script language=javascript>alert('你的江湖["& hygrade &"]级会员兑换完成！');history.back();</script>"
%>