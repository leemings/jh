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
sl=abs(int(Request.form("wpsl")))
if sl<1 or sl>50 then
	Response.Redirect "../../error.asp?id=72"
end if
action=lcase(trim(request.querystring("action")))
if InStr(action,"or")<>0 or InStr(action,"=")<>0 or InStr(action,"`")<>0 or InStr(action,"'")<>0 or InStr(action," ")<>0 or InStr(action,"　")<>0 or InStr(action,"'")<>0 or InStr(action,chr(34))<>0 or InStr(action,"\")<>0 or InStr(action,",")<>0 or InStr(action,"<")<>0 or InStr(action,">")<>0 then Response.Redirect "../../error.asp?id=54"
if action="" then 
		Response.Write "<script Language=Javascript>alert('提示：操作错误,指定物品名！！');location.href = 'javascript:history.go(-1)'</script>"
		response.end
end if
name=sjjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT 操作时间 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
sj=DateDiff("s",rs("操作时间"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('对不起系统限制，请等["&s&"秒钟]再操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if	
rs.close
rs.open "select * from w where c='药品' and a='" & action & "' and b=" & sjjh_id & " and i>="&sl,conn
	if rs.eof or rs.bof then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('提示：你的物品数量不足！');location.href = 'javascript:history.go(-1)'</script>"
			response.end
	end if
id=rs("ID")
nei=int(rs("e"))*sl
ti=int(rs("f"))*sl
conn.execute "update w set i=i-"& sl &" where id=" & id
conn.execute "delete * from w where i<=0"
conn.execute "update 用户 set 内力=内力+" & nei & ", 体力=体力+" & ti & ",操作时间=now() where 姓名='"&sjjh_name&"'"
mess="你吃用了药物"& sl &"个，增加内力" & nei & "，增加体力" & ti & "！"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=Application("sjjh_chatroomname")%></title></head>

<body background="back.gif" oncontextmenu=self.event.returnValue=false>
<table border="0" align="center" width="300" cellspacing="0" cellpadding="0">
<tr align="center">
<td width="300" height="30"><font
color="FF6600"><b>江湖提示</b></font></td>
</tr>
<tr align="center">
<td>
<table width="260">
<tr>
<td>
<p align="center" style="font-size:14;color:red"><%=mess%></p>
</td>
</tr>
</table>
</td>
</tr>
<tr align="center">
<td width="500" height="30"><a href="javascript:location.replace('eat.asp');" onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回装备一览</a></td>
</tr>
</table>

</body>