<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
sl=abs(int(Request.form("xlsl")))
if sl<1 or sl>50 then
	Response.Redirect "../../error.asp?id=72"
end if
action=request.querystring("action")
name=aqjh_name
if action<>"" then
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT 操作时间 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
sj=DateDiff("s",rs("操作时间"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('对不起系统限制，请等["&s&"秒钟]再操作！');}</script>"
	Response.End
end if	
rs.close
rs.open "select * from x where a='" & action & "' and b=" & aqjh_id & " and c>0",conn
if rs.eof or rs.bof then
mess="你没有这种物品！"
else
if int(rs("c"))<sl then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的物品只有"&rs("c")&"个，而你想服用"&sl&"个,你的物品不够！');</script>"
	%>
	<script language="vbscript">
	location.href = "javascript:history.go(-1)"
	</script>
	<%
	response.end
end if
id=rs("ID")
gx=rs("d")
jjsj=int(rs("e"))*sl
conn.execute "update x set c=c-"& sl &" where id=" & id
conn.execute "delete * from x where c<=0"
select case gx
case "杀人数"
	conn.execute "update 用户 set "& gx &"="& jjsj &",操作时间=now() where 姓名='"&aqjh_name&"'"
case "解除通缉"
	conn.execute "update 用户 set 门派='游侠',身份='弟子',道德=100 where 姓名='"&aqjh_name&"'"
case "运气"
	chance=array("道德","魅力","内力","体力","武功")
	randomize()
	rnd2=int(rnd*4)
	jjsj=int(rnd*1000)+500
	gx=chance(rnd2)
	conn.execute "update 用户 set "& gx &"="& gx &"+"& jjsj &" ,操作时间=now() where 姓名='"&aqjh_name&"'"
case else
	conn.execute "update 用户 set "& gx &"="& gx &"+"& jjsj &" ,操作时间=now() where 姓名='"&aqjh_name&"'"
end select
mess="你吃用了药物"& action & sl &"个，增加" & gx & jjsj & "点！"
end if
end if
tmprs=conn.execute("Select sum(g) As data from w where j=True and b="&aqjh_id)
zbgj=tmprs("data")
if isnull(zbgj) then zbgj=0
tmprs=conn.execute("Select sum(h) As data from w where j=True and b="&aqjh_id)
zbfy=tmprs("data")
if isnull(zbfy) then zbfy=0
djgy=aqjh_jhdj*1200+4500
djfy=aqjh_jhdj*1120+3000
conn.execute "update 用户 set 攻击="&djgy+zbgj&",防御="&djfy+zbfy &" where 姓名='"&aqjh_name&"'"
conn.execute "delete * from w where  i<=0"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=Application("aqjh_chatroomname")%></title></head>

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