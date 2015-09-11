<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');window.close();</script>"
	Response.End
end if
my=aqjh_name
money=abs(request.form("money"))
if money<>1000 and  money<>10000 and money<>100000 and money<>1000000 and money<>2000000 and money<>10000000 then 
	Response.Write "<script Language=Javascript>alert('你想作什么？！');location.href = 'javascript:history.back()';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=210"
	response.end
end if
sj=DateDiff("n",rs("操作时间"),now())
if sj<15 then
	s=15-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('对不起系统限制，请等["&s&"分钟]再操作！');}</script>"
	Response.End
end if	

if rs("银两")<money then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('您的银两不够呀，虽然是师徒，不过没有钱是万万不行的！');location.href = 'javascript:history.back()';;</script>"
	response.end
end if
conn.execute "update 用户 set 武功=武功+"&int(money/1000)*15&",银两=银两-"& money &",操作时间=now() where 姓名='"&aqjh_name&"'"
rs.close
conn.close
set rs=nothing
set conn=nothing
Response.Write "<script Language=Javascript>alert('你与师傅在密室中苦心学习！终于使自己的武功大进,武功：+"& ((money/1000)*15) &"点，花费银两：-"&money&"两！');location.href = 'javascript:history.back()';;</script>"
%>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../bg.gif">