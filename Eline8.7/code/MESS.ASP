<%@ LANGUAGE=VBScript codepage ="936" %>
<%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")%>
<html>
<head>
<title>江湖消息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#88AFD7">
<%
if sjjh_name<>"" then
rs.open "select 师傅,师傅交钱,银两 from 用户 where 姓名='" & sjjh_name & "'",conn
sf=rs("师傅")
Response.Write "<div align='center'><font size=-1>师傅："& sf &"</font></div>"
if rs("师傅")<>"" and rs("师傅")<>"无" and rs("师傅交钱")<>"交钱"&Day(date()) then
	yin=int(rs("银两")*0.05)
	conn.execute "update 用户 set 银两=银两-"& yin &",师傅交钱='交钱"& Day(date()) &"' where 姓名='"& sjjh_name &"'"
	conn.execute "update 用户 set 银两=银两+"& yin &" where 姓名='"& sf &"'"
	Response.Write "<div align='center'><font size=-1>"& sjjh_name &"今天上交师门"& yin &"两白银，来孝敬师傅："& sf &"</font></div>"
end if
rs.close
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r
rs.open "select b,a from k where d=false and a='"& sjjh_name & "'",conn
if not(rs.BOF or rs.EOF) then
	Response.Write "<br>你在江湖钱装有贷款，贷款日期："& rs("b") &"注意还款,7天不还删除ID！"
end if
rs.close
rs.open "select * from k where d=false and a='"& sjjh_name & "' and DateDiff('d',b,#" & sj & "#)>7",conn
if not(rs.BOF or rs.EOF) then
	name=rs("a")
	conn.Execute ("update k set d=true where a='"&name&"'")
	conn.Execute ("delete * from 用户 where 姓名='"&name&"'")
	'conn.Execute ("insert into l(b,a,c,e,d) values ('" & name & "',"&sqlstr(sj)&",'高利贷杀手','欠债还钱，没钱要你命','人命')")
	Response.Write "因为你7天没有还贷款，江湖ID被删除了！"
	sjjh_name=""
	Session.Abandon
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
else
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if
else%>
<font size=-1><%=Application("sjjh_chatroomname")%>欢迎您！</font> 
<%end if
%>
<div align="center"></div>
</body>
</html>
