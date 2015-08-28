
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
un=session("yx8_mhjh_username")
usergrade=session("yx8_mhjh_usergrade")
if un="" then Response.Redirect "../../error.asp?id=016"
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
        response.write "禁止站外提交！"
        response.end
end if
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hs")=0 then 
Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if 
id=LCase(trim(request.querystring("id")))
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 or InStr(id,"&")<>0 or InStr(id,"<")<>0 then 
Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');window.close();}</script>"
Response.End 
end if
if usergrade<10 then 
	Response.Write "<script Language=Javascript>alert('提示：你来这里作什么，想死呀！等级不够10');window.close();</script>"
	response.end
end if
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.open "SELECT 银两 FROM 用户 where 姓名='"&un&"' and 银两>100000",conn
if rst.bof or rst.eof then
	rst.close
	conn.close
	set conn=nothing
	set rst=nothing
	Response.Write "<script language=JavaScript>{alert('需要10万两！');location.href = 'javascript:history.back()';}</script>"
	response.end
end if
rst.close
%><!--#include file="data.asp"--><%
rs.open "select id from mj where id="& id &" and 所有者='"&un&"' and 类型='轩'",connb
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	connb.close
	set connb=nothing
	Response.Write "<script Language=Javascript>alert('提示：你没有这招！');history.go(-1);</script>"
	response.end
end if
connb.execute "delete * from mj where id="& id &" and 所有者='"&un&"' and 类型='轩'"
conn.execute "update 用户 set 银两=银两-100000 where 姓名='"&un&"'"
rs.close
set rs=nothing
connb.close
set connb=nothing
set rst=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：删除轩辕武功成功,银两减少10万');location.href = 'javascript:history.go(-1)';</script>"
%>
