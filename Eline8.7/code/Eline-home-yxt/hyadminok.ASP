<%Response.Expires=0
Response.Buffer=true
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
username=Request.Form("search")
sjyear=Request.form("year")
sjmonth=Request.form("month")
sjday=Request.form("day")
hygrade=int(Request.form("hygrade"))
hydate="#"&sjyear&"/"&sjmonth&"/"&sjday&"#"
select case hygrade
case 1
jhgrade=1
hyvalue=31250
hyjia=0
case 2
jhgrade=1
hyvalue=90000
hyjia=100
case 3
jhgrade=5
hyvalue=250000
hyjia=200
case 4
jhgrade=7
hyvalue=490000
hyjia=300
end select
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("sjjh_usermdb")
Set rs=Server.CreateObject("ADODB.RecordSet")
sqlstr="SELECT * FROM 用户 where 姓名='"&username&"'"
rs.Open sqlstr,conn
if rs.EOF or rs.BOF then
	Response.Write "<script language=javascript>alert('抱歉你所要查找的人我们找不到！请查看是否正确！');history.back();</script>"
else
	if rs("会员等级")=0 and rs("grade")<>10 then
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','管理记录','[增加会员]姓名："&username&"，等级："&hygrade&" 日期："&hydate&"')"
		sql="update 用户 set 会员等级="& hygrade &",会员日期="& hydate &",内力加=内力加+"& hyjia &",体力加=体力加+"& hyjia &",武功加=武功加+"& hyjia &",grade="& jhgrade &",allvalue=allvalue+"& hyvalue &" where 姓名='"&username&"'"
		Set rs=conn.Execute(sql)
	else
		Response.Write "<script language=javascript>alert('抱歉["&username&"]现在已经是["&rs("会员等级")&"级]江湖会员，请手动更改！');history.back();</script>"
	end if



Response.Write "<script language=javascript>alert('恭喜你["&username&"]的江湖["& hygrade &"]会员设置完成！');history.back();</script>"
%>