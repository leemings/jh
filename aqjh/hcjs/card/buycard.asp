<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<!--#include file="cardset.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
id=lcase(trim(request("id")))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"　")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../error.asp?id=54"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 姓名 from 用户 where 姓名='"&aqjh_name&"'",conn,3,3
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你不是江湖中人，滚！');}</script>"
	Response.End
end if
rs.close
rs.open "SELECT * FROM b where ID=" & id & " and h>0 and b='卡片'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('没有你要购买的会员卡片？!');location.href = 'card.asp';}</script>"
	response.end
end if
card=rs("a")
hyyin=rs("h")
sm=rs("c")
rs.close
rs.open "select 会员金卡,w5 from 用户 where  姓名='"&aqjh_name&"'",conn,3,3
if hyyin>rs("会员金卡") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你的会员金卡不足！!');location.href = 'card.asp';}</script>"
	response.end
end if
mycard=add(rs("w5"),card,1)
sl=mywpsl(rs("w5"),card)+1
if sl>max then
   Response.Write "<script language=JavaScript>{alert('你的卡片["&card&"]已经超出20个！请等一下再购买！');window.close();}</script>"
else
conn.execute "update 用户 set w5='"&mycard&"',会员金卡=会员金卡-" &hyyin & " where  姓名='"&aqjh_name&"'"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的卡片["&card&"]购买成功,现有"&sl&"个！');location.href = 'card.asp';}</script>"
end if%>