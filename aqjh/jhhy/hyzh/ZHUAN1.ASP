<body bgcolor="#000000">
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
username=Session("aqjh_name")
hymonth=1
hygrade=1
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	conn.open Application("aqjh_usermdb")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	sqlstr="SELECT * FROM 用户 where 姓名='"&username&"'"
	rs.Open sqlstr,conn,1,2
	if rs.EOF or rs.BOF then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('江湖提示：无此用户！请查看是否正确！');history.back();</script>"
		Response.End
	end if
        if rs("会员金卡")<500 then
		Response.Write "<script language=javascript>alert('江湖提示：抱歉你的条件不够，金卡500元不够');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
	    if rs("金币")<1000 then
		Response.Write "<script language=javascript>alert('江湖提示：抱歉你的条件不够，金币1000个不够');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
	    if rs("银两")<80000000 then
		Response.Write "<script language=javascript>alert('江湖提示：抱歉你的条件不够，银子8千万不够');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
		if rs("会员等级")<>0  then
		Response.Write "<script language=javascript>alert('江湖提示：抱歉你已经是["&rs("会员等级")&"级]江湖会员！');history.back();</script>"
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
	jy=125000
	
	sj=n & "-" & y & "-" & r
	
	rs("会员等级")=hygrade
	rs("会员日期")=cdate(sj)
	rs("allvalue")=rs("allvalue")+jy
	rs.Update
	conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','操作','[兑换"&hygrade&"级会员]日期："&sj&"')"
	conn.execute "update 用户 set 会员金卡=会员金卡-500,金币=金币-1000,银两=银两-80000000 where 姓名='"&username&"'"
       rs.close
set rs=nothing
	Response.Write "<script language=javascript>alert('你的江湖["& hygrade &"]级会员兑换完成！');history.back();</script>"
conn.close
set conn=nothing
says="官方消息</font>快乐江湖大侠"& aqjh_name &"转换1级别会员成功,非会员的朋友赶快到会员转换处,去转换会员吧!"&mess			'聊天数据
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session("nowinroom") & ");<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%>
