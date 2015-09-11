<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if aqjh_grade<10 or aqjh_name<>application("aqjh_user") then
 Response.Write "<script Language=JavaScript>alert('提示：你没有权限设置奖品');window.close();</script>"
 Response.End
end if
%>
<!--#include file="data.asp"-->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
sql1="delete * from 养猪风云榜"
sql2="delete * from 养猪投票者"
connt.execute(sql1)
connt.execute(sql2)
 Response.Write "<script Language=JavaScript>alert('提示：已经将所有数据清理，以备下一次大赛的开始！');window.location.href ='index.asp';</script>"
 Response.End
%>