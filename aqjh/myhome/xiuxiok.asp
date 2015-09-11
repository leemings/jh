<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
act=clng(trim(request("act")))
if act<>1 and act<>2 then act=1
if act=1 then
	sj=clng(trim(Request("sj")))
	if sj<1 or sj>72 then
		Response.Write "<script Language=Javascript>alert('提示：时间选择错误!');location.href = 'javascript:history.go(-1)';</script>"
		response.end
	end if
else
	if Hour(time())<21 or Hour(time())>23 then
		Response.Write "<script Language=Javascript>alert('提示：已婚侠客休息时间为每天21-23点,每次8小时!');location.href = 'javascript:history.go(-1)';</script>"
		response.end
	end if
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.Open "select h_耐久度 from house  where "& session("myroom") &"='"& aqjh_name &"'", conn, 1, 1
if rs.Eof and rs.Bof then
	rs.Close
	Set rs = Nothing
	conn.Close
	Set conn = Nothing
   	Response.Write "<script language=JavaScript>{alert('提示:没有购买房子是不能休息的！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if rs("h_耐久度")<10 then
	rs.Close
	Set rs = Nothing
	conn.Close
	Set conn = Nothing
   	Response.Write "<script language=JavaScript>{alert('提示:你的屋子损坏严重,请修复再来吧！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
rs.close
if act=1 then
	rs.open "SELECT 状态,登录,事件原因,体力 from [用户] WHERE 姓名='" & aqjh_name & "'",conn,1,3
	if rs("状态")<>"正常" then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：你现在的状态不正常不能进行休息!');location.href = 'javascript:history.go(-1)';</script>"
		response.end
	end if
	rs("状态")="眠"
	rs("登录")=DateAdd("h", sj, now())
	rs("事件原因")="您在自己家中卧室中休息了:"& sj &"小时,上线时间是在:"& DateAdd("h", sj, now()) 
	rs("体力")=rs("体力")+sj*80000
	rs.update
	rs.close
	show="提示：休息成功,增加体力:"& sj *80000 &"点\n,按确定重退出江湖,请在"& sj &"小时后登录江湖!"
else
	rs.open "SELECT 配偶 from [用户] WHERE 姓名='" & aqjh_name & "'",conn,1,1
	mywife=rs("配偶")
	rs.close
	rs.open "SELECT 状态 from [用户] WHERE 姓名='" & mywife & "'",conn,1,1
	if rs.Eof or rs.Bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：找不到你的配偶资料!');location.href = 'javascript:history.go(-1)';</script>"
		response.end
	end if
	wifezt=rs("状态")
	rs.close
	rs.open "SELECT 配偶,状态,登录,事件原因,体力 from [用户] WHERE 姓名='" & aqjh_name & "'",conn,1,3
	rs("状态")="眠"
	rs("登录")=DateAdd("h", 8, now())
	rs("事件原因")="天已经晚了,是回家休息的时候了,每天保证8小时休息对身体好,体力增加100万点!上线时间是在:"& DateAdd("h", 8, now())
	rs("体力")=rs("体力")+15000
	rs.update
	rs.close
	'这里以后要判断小孩子功能,以确定怀孕.
	show="提示：已婚大侠休息成功,增加体力:100万点\n,按确定重退出江湖,请在8小时后登录江湖!"
end if
rs.Open "select h_耐久度 from house  where "& session("myroom") &"='"& aqjh_name &"'", conn, 1, 3
rs("h_耐久度")=rs("h_耐久度")-1
rs.update
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('"& show &"');top.location.href = '../exit.asp';</script>"
response.end
%>