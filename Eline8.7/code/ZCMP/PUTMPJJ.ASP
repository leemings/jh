<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<title>门派基金♀wWw.51eline.com♀</title>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../bgcheetah.gif">
<%if sjjh_name="" then Response.Redirect "../error.asp?id=210"
my=sjjh_name
money=abs(request.form("money"))
if money<>1000 and  money<>10000 and money<>100000 and money<>1000000 then
 	Response.Write "<script language=JavaScript>{alert('严重警告，不要搞乱');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * from 用户 where 姓名='"&sjjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=210"
	response.end
end if
if trim(rs("门派"))="游侠" or trim(rs("门派"))="无" or trim(rs("门派"))="未知"  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你是游侠，还没有自己的门派，你想搞什么呀？');location.href = 'javascript:history.back()';</script>"
	response.end
end if
sj=DateDiff("n",rs("操作时间"),now())
if sj<5 then
	s=5-sj
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
	Response.Write "<script Language=Javascript>alert('你的钱好象不够呀！');location.href = 'javascript:history.back()';</script>"
	response.end
end if
mp=rs("门派")
conn.execute "update 用户 set 银两=银两-"& money &",道德=道德+"&int(money/500)&",门派基金=门派基金+"&money &",操作时间=now() where 姓名='"&sjjh_name&"'"
conn.execute "update p set h=h+"& money &" where a='" & mp &"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('您向你的门派中捐献了"& money &"两，道德上涨："&int(money/500)&"点帮中的弟子对你感激不尽！');location.href = 'javascript:history.back()';</script>"
%>
