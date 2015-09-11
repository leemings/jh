<!--#include file="../showchat.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
name=trim(request("myid"))
if name="" then
 Response.Write "<script Language=JavaScript>alert('提示：出错!');window.close();</script>"
 Response.End
end if
if name<>aqjh_name then
 Response.Write "<script Language=JavaScript>alert('提示：捣什么乱啊，又不是你中奖!');window.close();</script>"
 Response.End
end if
%>
<!--#include file="data.asp"-->
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "Select * from 猪圈设置",connt,3,3
jz=rs("投票结束")
if cdate(jz)>now() then
 Response.Write "<script Language=JavaScript>alert('提示：还没有开奖呢？你急什么呀？');window.close();</script>"
 Response.End
end if
rs.close
sql="Select * from 养猪风云榜 where 参赛人='"&name&"'"
rs.open sql,connt,1,1
if rs.recordcount=0 then
 Response.Write "<script Language=JavaScript>alert('提示：此人并没有参赛啊！');window.close();</script>"
 Response.End
else
   hjsj=rs("获奖时间")
   if rs("领奖")=true then
 Response.Write "<script Language=JavaScript>alert('提示：你搞什么啊，你已经领过奖了！');window.close();</script>"
 Response.End
   end if
   if rs("名次")<1 or rs("名次")>3 then
 Response.Write "<script Language=JavaScript>alert('提示：你又没中奖，来干什么啊？');window.close();</script>"
 Response.End
   end if
   sj=DateDiff("d",hjsj,now())
   if sj>15 then
 Response.Write "<script Language=JavaScript>alert('提示：你的奖品已经过期！');window.close();</script>"
 Response.End
   end if
  select case rs("名次")
  case 1
    sql="update 用户 set 状元=状元+100,金币=金币+500,会员金卡=会员金卡+10 where 姓名='"&name&"'"
    mess="荣获养猪大赛<font color=red>一等奖</font>，官府奖励金币500个，现金(会员金卡)10元，站长亲自颁奖！恭喜了！"
  case 2
    sql="update 用户 set 状元=状元+50,金币=金币+300,会员金卡=会员金卡+5 where 姓名='"&name&"'"
    mess="荣获养猪大赛<font color=red>二等奖</font>，官府奖励金币300个，现金(会员金卡)5元，站长亲自颁奖！恭喜了！"
  case 3
    sql="update 用户 set 状元=状元+20,金币=金币+100,会员金卡=会员金卡+3 where 姓名='"&name&"'"
    mess="荣获养猪大赛<font color=red>三等奖</font>，官府奖励金币100个，现金(会员金卡)3元，站长亲自颁奖！恭喜了！"
  case else
     Response.Write "<script Language=JavaScript>alert('提示：奖品出错！');window.close();</script>"
     Response.End
  end select
  sql2="update 养猪风云榜 set 领奖=true where 参赛人='"&name&"'"
  conn.execute(sql)
  connt.execute(sql2)
end if
rs.close
sql="Select zhu from 用户 where 姓名='"&aqjh_name&"'"
rs.open sql,conn,1,1
myhua=trim(rs("zhu"))
fhua=split(myhua,"|")
huart="0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|"
if myhua<>"" then
seedsm=int(fhua(0))
if seedsm>0 then
huasj=fhua(1)
myhua_ok=seedsm&"|"&huasj&"|"&huart
else
myhua_ok="3|"&now()&"|"&huart
end if
else
myhua_ok="3|"&now()&"|"&huart
end if
conn.execute "update 用户 set zhu='"&myhua_ok&"' where 姓名='"&aqjh_name&"'"
call showchat ("<bgsound src=../chat/wav/hkjr.mp3 loop=2><font color=red>【养猪大赛】</font><img src=pic/hk.gif>"&aqjh_name&mess)
Response.Write "<script Language=JavaScript>alert('提示：恭喜你已经领取了养猪大赛奖品');window.close();</script>"
Response.End
%>