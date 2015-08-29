<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
id=trim(request("id"))
if instr(id,"官")<>0 then
		Response.Write "<script language=JavaScript>{alert('严重警告，不要搞乱');window.close();}</script>"
		Response.End
end if
my=trim(request("my"))
if my<>sjjh_name then
		Response.Write "<script language=JavaScript>{alert('严重警告，不要搞乱');window.close();}</script>"
		Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM 用户 WHERE 身份<>'掌门' and 姓名='"&sjjh_name&"'",conn
if rs.eof or rs.bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('提示：江湖上找不到你的资料或者你是掌门！');window.close();</script>"
		response.end
end if
	if rs("门派")="" or rs("门派")="游侠" or rs("门派")="无" then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('提示：你并无门派！！');window.close();</script>"
		response.end
	end if
	if rs("门派")="出家" or rs("门派")="天网"  then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('提示：你是出家人，或是天网杀手不可以在这里离开!！！');window.close();</script>"
		response.end
	end if
	if rs("银两")<20000000 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('提示：离开门派需要2000万银两手续费，你有吗？！');window.close();</script>"
		response.end
	end if

if rs("会员等级")<6 then
	message="你离开了" & rs("门派") & "，损失内力、武功、体力各5000，银两及存款均降到1000两！"
	if rs("存款")>1000 then
		conn.execute "update 用户 set 门派='游侠',身份='弟子', 内力=内力-5000,体力=体力-5000,武功=武功-5000,grade=1,银两=1000,存款=1000 where 姓名='"&sjjh_name&"'"
	else
		conn.execute "update 用户 set 门派='游侠',身份='弟子', 内力=内力-5000,体力=体力-5000,武功=武功-5000,grade=1,银两=0 where 姓名='"&sjjh_name&"'"
	end if
else
	message="你离开了" & rs("门派") & "，因为你是江湖贵宾，所以原有内力，金钱不变！"	
	conn.execute "update 用户 set 门派='游侠',身份='弟子',grade=1 where 姓名='"&sjjh_name&"'"
end if
	conn.execute "update p set c=c-1 where a='" & id & "'"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>〖<%=Application("sjjh_chatroomname")%>〗</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body background="../jhimg/bk_hc3w.gif" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"><font size="+3" color="#FF0000">离开门派</font><br>
<br>
<br>
<%=Application("sjjh_chatroomname")%>告示：<%=message%> </div>
<hr>
<div align="center">〖<%=Application("sjjh_chatroomname")%>〗<a href="mailto:eline_email@etang.com.com"><%=Application("sjjh_chatroomname")%></a> </div>
</body>
</html>
