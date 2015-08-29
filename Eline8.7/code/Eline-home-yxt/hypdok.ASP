<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
username=Request.Form("search")
hymonth=Request.form("hymonth")
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
		Response.Write "<script language=javascript>alert('抱歉你所要查找的人我们找不到！请查看是否正确！');history.back();</script>"
		Response.End
	end if
	if rs("会员")=true then
		hydate=rs("会员结束")
	else
		hydate=date()
	end if
	n=Year(hydate)
	y=Month(hydate)
	r=Day(hydate)
	y=y+hymonth
	if y>12 then 
		y=y-12
		n=n+1
	end if
	sj=n & "-" & y & "-" & r
	rs("会员")=true
	rs("会员结束")=cdate(sj)
	rs.Update
	conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','管理记录','[增加泡点制会员]姓名："&username&"，日期："&sj&"')"
	set rs=nothing
	conn.Close
	set conn=nothing
	Response.Write "<script language=javascript>alert('恭喜你["&username&"]的江湖[泡点制]会员设置完成！');history.back();</script>"
%>