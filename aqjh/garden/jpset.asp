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
name=request("name")
j=request("j")
if name="" or j="" then
 Response.Write "<script Language=JavaScript>alert('提示：出错!');window.close();</script>"
 Response.End
end if
%>
<!--#include file="data.asp"-->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.open "Select * from 花园设置",connt,3,3
jz=rs("投票结束")
if cdate(jz)>now() then
 Response.Write "<script Language=JavaScript>alert('提示：投票没结束，不能颁奖！');window.close();</script>"
 Response.End
end if
rs.close
sql="Select * from 花园风云榜 where 参赛人='"&name&"'"
rs.open sql,connt,1,1
if rs.recordcount=0 then
 Response.Write "<script Language=JavaScript>alert('提示：此人并没有参赛啊！');window.close();</script>"
 Response.End
else
   if rs("名次")>0 and j<>0 then
       Response.Write "<script Language=JavaScript>alert('提示：不能第二次颁奖！');window.close();</script>"
       Response.End
   else
   select case j
   case 0
    sql2="update 花园风云榜 set 名次=0 where 参赛人='"&name&"'"
    connt.execute(sql2)
    Response.Write "<script Language=JavaScript>alert('提示："&name&"名次已经清除！');window.close();</script>"
    Response.End
   case 1,2,3
    sql2="update 花园风云榜 set 名次="&j&",获奖时间='"&now()&"' where 参赛人='"&name&"'"
    connt.execute(sql2)
    Response.Write "<script Language=JavaScript>alert('颁奖："&name&"已经设置为"&j&"等奖！');window.close();</script>"
    Response.End
   case else
       Response.Write "<script Language=JavaScript>alert('提示：出错！');window.close();</script>"
       Response.End
   end select
   end if
end if
%>