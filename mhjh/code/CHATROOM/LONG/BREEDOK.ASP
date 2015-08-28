<%
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"chatroom")=0 then 
Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if 
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
        response.write "禁止站外提交！"
        response.end
end if
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
id=Request.QueryString ("id")
if id="" then Response.end
if not isnumeric(id) then Response.End 
animalname=Request.QueryString ("animalname")
if animalname="" then Response.End 
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
rs.Open "select 体力,id,数量 from 物品 where 所有者='"&username&"' and 数量>0 and id="&id&"",conn
if rs.BOF or rs.EOF then
Response.Write "<script language=JavaScript>{alert('  对不起！\n  您没有此种龙粮，不能喂养神龙！！！\n  按 [确定] 返 回！');parent.history.go(-1);}</script>"
rs.Close
set rs=nothing
conn.Close 
set conn=nothing
Response.End 
end if
c4=rs("体力")
id=rs("id")
c9=rs("数量")
rs.Close 
%><!--#include file="data.asp"--><% 
sql= "select * from myanimal where username='"&username&"' and rest=0 and animalname='"&animalname&"'"
rs.open sql,connb,1,1  
if rs.BOF or rs.eof then
Response.Write "<script language=JavaScript>{alert('  对不起！\n  您没有此种神龙！！！\n  按 [确定] 返 回！');parent.history.go(-1);}</script>"
rs.Close 
set rs=nothing
conn.Close 
set conn=nothing
connb.Close 
set connb=nothing
Response.End 
end if
id1=rs("id")
attack=rs("attack")
allattack=rs("allattack")
if attack>=allattack then
rs.Close
set rs=nothing
conn.Close
set conn=nothing
connb.Close 
set connb=nothing
Response.Write "<script language=JavaScript>{alert('  对不起！\n  您的神龙目前己达到最佳状态，不需要再补了！！！\n  按 [确定] 返 回！');parent.history.go(-1);}</script>"
Response.End 
end if

if c9>=1 then 
conn.Execute ("update 物品 set 数量=数量-1 where 所有者='"&username&"' and id="&id)
else
conn.Execute "delete * from 物品 where 所有者='"&username&"' and id="&id&""
end if
conn.Execute "delete * from 物品 where 所有者='"&username&"' and 属性='龙粮' and 数量=0"
connb.Execute ("update myanimal set attack=attack+"&c4&"  where id="&id1)
rs.Close
set rs=nothing
conn.Close 
set conn=nothing
connb.Close 
set connb=nothing
Response.Write "<script language=JavaScript>{alert('  恭喜您！\n  你的神龙吃了你提供的龙粮！！！\n  按 [确定] 返 回！');location.href='myanimal.asp';}</script>"
%>