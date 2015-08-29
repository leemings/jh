<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * from 用户 where 姓名='"&sjjh_name&"'",conn
mp=rs("门派")
to1=rs("姓名")
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
	Response.Write "<script Language=Javascript>alert('你是游侠，还没有自己的门派，你想搞什么呀？');window.close();</script>"
	response.end
end if
if trim(rs("门派"))="官府"  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你是官员，不能申请！');window.close();</script>"
	response.end
end if
tmprs=conn.execute("Select count(*) As 数量 from 用户 where grade=4 and 身份='长老' and 门派='"& mp &"'")
	musers=tmprs("数量")
	set tmprs=nothing
	if isnull(musers) then musers=0
	if musers>=8 then
		Response.Write "<script language=JavaScript>{alert('["&sjjh_name&"]现在你门派的长老有8个了，不能再申请了！');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
if rs("身份")="长老" or rs("身份")="掌门" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你的身份以在这以上了，不用在申请！');window.close();</script>"
	response.end
end if
if rs("会员")=true then
conn.execute "update 用户 set 身份='长老',grade=4 where 姓名='" & to1 &"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你是会员，不用任何条件，就成为了门派的长老，当会员的滋味真好！');window.close();</script>"
	response.end
end if
dengji=rs("等级")
jj=rs("门派基金")
if dengji<25 or jj<30000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你的等级或门派基金还不够所需要求，申请失败，请继续努力！');window.close();</script>"
	response.end
end if
conn.execute "update 用户 set 身份='长老',grade=4 where 姓名='" & to1 &"'"
Response.Write "<script Language=Javascript>alert('恭喜你，你的长老身份申请成功，好好维护你的门派哦!');window.close();</script>"
response.end

conn.close
set rs=nothing
%>