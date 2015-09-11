<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
username=Request.Form("search")
hymonth=Request.form("hymonth")
hygrade=int(Request.form("hygrade"))
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
		Response.Write "<script language=javascript>alert('抱歉你所要查找的人我们找不到！请查看是否正确！');history.back();</script>"
		Response.End
	end if
	'if rs("会员等级")<>0 or rs("grade")=10 then
	'	Response.Write "<script language=javascript>alert('抱歉["&username&"]现在已经是["&rs("会员等级")&"级]江湖会员，请手动更改！');history.back();</script>"
	'	rs.close
	'	set rs=nothing
	'	conn.close
	'	set conn=nothing
	'	Response.End	
	'end if
	if rs("会员等级")>1 then
		hydate=rs("会员日期")
	else
		hydate=date()
	end if
	oldhy=rs("会员等级")
	n=Year(hydate)
	y=Month(hydate)
	r=Day(hydate)
	y=y+hymonth
	if y>12 then 
		y=y-12
		n=n+1
	end if
	jy=0
	if hygrade=1 then jy=31250
	if hygrade=2 then jy=90000
	if hygrade=3 then jy=250000
	if hygrade=4 then jy=490000
	if hygrade=5 then jy=1000000
	if hygrade=6 then jy=1800000
	if hygrade=7 then jy=2000000
	if hygrade=8 then jy=4500000
	if oldhy=1 then oldjy=31250
	if oldhy=2 then oldjy=90000
	if oldhy=3 then oldjy=250000
	if oldhy=4 then oldjy=490000
	if oldhy=5 then oldjy=1200000
	if oldhy=6 then oldjy=1800000
	if oldhy=7 then oldjy=2000000
	if oldhy=8 then oldjy=4500000
	sj=n & "-" & y & "-" & r
	if hygrade=8 then
                if username=Application("aqjh_user") then
                   response.write "<script language=javascript>alert('你操作的对象是江湖10级管理员,请慎用此操作！');history.back();</script>"
                   response.end
                end if
		rs("门派")="官府"
		rs("grade")=7
	end if
	rs("会员等级")=hygrade
	rs("会员日期")=cdate(sj)
	rs("allvalue")=rs("allvalue")+jy-oldjy
	rs.Update
	conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','管理记录','[增加等级制会员]姓名："&username&"，等级："&hygrade&" 日期："&sj&"')"
	set rs=nothing
	conn.Close
	set conn=nothing
	Response.Write "<script language=javascript>alert('恭喜你["&username&"]的江湖["& hygrade &"]会员设置完成！');history.back();</script>"
%>