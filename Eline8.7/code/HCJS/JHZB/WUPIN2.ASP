<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
name=lcase(trim(request("name")))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")

rs.open "select b,f,g,j,k FROM b WHERE a='" & name &"'",conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：此类物品在数据库中不存在！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
nai=rs("j")
lb=rs("b")
wpfy=rs("g")
wpgj=rs("f")
wptx=rs("k")
select case lb
	case "头盔"
		lb="z1"
	case "装饰"
		lb="z2"
	case "盔甲"
		lb="z3"
	case "左手"
		lb="z4"
	case "右手","锻造"
		lb="z5"
	case "双脚"
		lb="z6"
	case else
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：类别数据出错，与程序开发商联系！');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End
end Select
rs.close

rs.open "SELECT 操作时间,w3,攻击,防御,"&lb&" FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,3,3
sj=DateDiff("s",rs("操作时间"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('对不起系统限制，请等["&s&"秒钟]再操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end If
oldgj = rs("攻击")
oldfy = rs("防御")
oldz = rs(lb)
'Response.Write rs("w3")

duyao=abate(rs("w3"),name,1)
'Response.Write "===="&duyao
rs.close

'删除掉已经传的装备
if Not isnull(oldz) and oldz<>"" then 
	haveWareArr = Split(oldz,"|")
	ArrLen=UBound(haveWareArr)
'	Response.Write "---------"&ArrLen
	If ArrLen >= 4 Then
'		Response.Write "--------11-"&ArrLen
		oldgj = oldgj - haveWareArr(2)
		oldfy = oldfy - haveWareArr(3)
		If isnull(duyao) Then
			duyao = ""
		End if	
		duyao = duyao & haveWareArr(0) &"|1"& ";" 
		'如果存在多件相同的装备，还会有问题的
	End if
End if

conn.execute "update 用户 set w3='"&duyao&"',"&lb&"='"&name&"|"&nai&"|"&wpgj&"|"&wpfy&"|"&wptx&"',攻击="&(oldgj+wpgj)&",防御="&(oldfy+wpfy)&" where  姓名='"&sjjh_name&"'"
'Response.Write "update 用户 set w3='"&duyao&"',"&lb&"='"&name&"|"&nai&"|"&wpgj&"|"&wpfy&"|"&wptx&"',攻击="&(oldgj+wpgj)&",防御="&(oldfy+wpfy)&" where  姓名='"&sjjh_name&"'"
'update 用户 set w3='',z1='黄金帽|70|100|750|无' where 姓名='回首当年'
'return
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：装备物品["&name&"]完成\n按确定返回');location.href = 'javascript:history.go(-1)';</script>"
response.end
%>