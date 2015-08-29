<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/jhjs")=0 then 
Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if 
id=request("id")
if not isnumeric(id) then Response.Redirect "../../error.asp?id=54"
'sl=int(abs(Request.form("sl")))
sl=1
if sl<1 or sl>9 then
	Response.Redirect "../../error.asp?id=72"
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM b where ID=" & id,conn
if rs("b")<>"右手" and rs("b")<>"左手" and rs("b")<>"盔甲" and rs("b")<>"头盔" and rs("b")<>"双脚" and rs("b")<>"装饰" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=72"
	response.end
end if
wu=rs("a")
yin=rs("h")
lx=rs("b")
gj=rs("f")
fy=rs("g")
sm=rs("b")
rs.close
rs.open "select 会员等级,银两,操作时间 from 用户 where id="&sjjh_id,conn
hy=rs("会员等级")
if hy>0 then
	yin=int(yin/2)
end if

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
if yin*sl>rs("银两") then
	Response.Write "<script Language=Javascript>alert('提示：购买不成功，原因：你的银两不够了');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
conn.execute "update 用户 set 银两=银两-" & yin*sl & ",操作时间=now() where id="&sjjh_id
rs.close
'开始保存
rs.open "select id from w where a='"& wu&"' and b="&sjjh_id
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into w (a,b,c,g,h,i,l,d) values ('"&wu&"',"& sjjh_id&",'"& lx&"',"& gj &","& fy &","&sl&","&int(yin)&",'"&sm&"')"
else
	id=rs("id")
	conn.execute "update w set l=" & int(yin) & ",i=i+" & sl & ",d='"&sm&"' where id="&id
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<body topmargin="6" leftmargin="0" bgcolor="#FFFFFF" background="../JHZB/bg.gif">
<%
if hy>0 then
	Response.Write "<script Language=Javascript>alert('提示：会员购买物品打5折,购买"&wu&sl&"个成功！');location.href = 'javascript:history.go(-1)';</script>"
else
	Response.Write "<script Language=Javascript>alert('提示：购买"&wu&sl&"个成功,如果您是会员可以打5折！');location.href = 'javascript:history.go(-1)';</script>"
end if
%>
