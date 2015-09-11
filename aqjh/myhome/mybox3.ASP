<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if isonline(aqjh_name,session("inroom"))=false then
	Response.Write "<script Language=Javascript>alert('提示：使用本功能必须要进入聊天室！');</script>"
	Response.End
end if
name=trim(Request.form("name"))
wpsl=clng(trim(Request.form("wpsl")))
lb=trim(Request.form("lb"))
pass=trim(Request.form("pass"))
Submit=trim(Request.form("Submit"))
if Submit<>"取出" and Submit<>"取出" then Submit="放入"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.Open "select * from house  where "& session("myroom") &"='"& session("aqjh_name") &"'", conn, 1, 1
if rs.Eof and rs.Bof then
	rs.Close
	Set rs = Nothing
	conn.Close
	Set conn = Nothing
   	Response.Write "<script language=JavaScript>{alert('提示:没有购买房子是不能使用储物箱的！');}</script>"
	Response.End
end if
roomjb=rs("h_级别")-1
if rs("h_耐久度")<10 then
	rs.Close
	Set rs = Nothing
	conn.Close
	Set conn = Nothing
   	Response.Write "<script language=JavaScript>{alert('提示:你的屋子损坏严重,请修复再来吧！');}</script>"
	Response.End
end if
rs.close
rs.open "SELECT * from mybox WHERE mb_账号='" & session("aqjh_name") & "'",conn,1,1
records=rs.recordcount
rs.close
if Submit="放入" then
	rs.open "SELECT "& lb &",会员等级 from 用户 WHERE 姓名='" & session("aqjh_name") & "'",conn,1,3
	if rs.Eof and rs.Bof then
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
    	Response.Write "<script language=JavaScript>{alert('提示:你的账号在江湖并不存在！');}</script>"
		Response.End
	end if
	hydj=rs("会员等级")
	nowsl=8+hydj*4+roomjb*4
	if records>=nowsl then
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
    	Response.Write "<script language=JavaScript>{alert('提示:您是"& hydj &"会员,最大存储数量:"& nowsl &"种！');}</script>"
		Response.End
	end if
	if lb="会员金卡" or lb="金币" then
		if rs(lb)<wpsl then
			rs.Close
			Set rs = Nothing
			conn.Close
			Set conn = Nothing
    		Response.Write "<script language=JavaScript>{alert('提示:["& lb &"]没有这么多不能操作！');}</script>"
			Response.End
		end if
		rs(lb)=rs(lb)-wpsl
	else
		if mywpsl(rs(lb),name)<wpsl then
  			rs.Close
    		Set rs = Nothing
   			conn.Close
  			Set conn = Nothing
   			Response.Write "<script Language=Javascript>alert('提示：物品"& name &"数量不足？');</script>"
			response.end
		end if
		rs(lb)=abate(rs(lb),name,wpsl)
	end if
	rs.update
	rs.close
	rs.open "SELECT * from mybox WHERE mb_账号='" & session("aqjh_name") & "' and mb_物品名='"& name &"'",conn,1,3
	if rs.Eof and rs.Bof then
		rs.addnew
		rs("mb_数量")=wpsl
	else
		rs("mb_数量")=rs("mb_数量")+wpsl
	end if
	rs("mb_账号")=trim(session("aqjh_name"))
	rs("mb_物品名")=name
	rs("mb_类别")=lb
	rs.update
else
	rs.open "SELECT * from mybox WHERE mb_账号='" & session("aqjh_name") & "' and mb_物品名='"& name &"'",conn,1,3
	if rs.Eof and rs.Bof then
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
    	Response.Write "<script language=JavaScript>{alert('提示:物品["& name &"]在储物箱中并不存在！');}</script>"
		Response.End
	end if
	if rs("mb_数量")<wpsl then
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
    	Response.Write "<script language=JavaScript>{alert('提示:物品["& name &"]数量不足,请重新输入！');}</script>"
		Response.End
	end if
	if rs("mb_数量")-wpsl>0 then
		rs("mb_数量")=rs("mb_数量")-wpsl
	else
		rs.delete
	end if
	rs.update
	rs.close
	rs.open "SELECT "& lb &",会员等级 from 用户 WHERE 姓名='" & session("aqjh_name") & "'",conn,1,3
	if lb="会员金卡" or lb="金币" then
		rs(lb)=rs(lb)+wpsl
	else
		zstemp=add(rs(lb),name,wpsl)
		if zstemp="1" then
			rs.Close
			Set rs = Nothing
			conn.Close
			Set conn = Nothing
			Response.Write "<script Language=Javascript>alert('提示：你的物品超过25种不能保存？');</script>"
			response.end
		end if
		rs(lb)=zstemp
	end if
	rs.update
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('提示:["& name &"]"& Submit &"成功!');parent.cw.location.reload();parent.wp.location.reload();}</script>"
Response.End
%>
