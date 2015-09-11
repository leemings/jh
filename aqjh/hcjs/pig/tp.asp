<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT id FROM 用户 WHERE  姓名='" & aqjh_name &"'",conn,2,2
aqjh_id=rs("id")
name=trim(request("name"))
if name=aqjh_name then
 Response.Write "<script Language=JavaScript>alert('提示：为防作弊，不得给自己投票！');window.close();</script>"
 Response.End
end if
rs.close
%>
<!--#include file="data.asp"-->
<%
rs.open "Select * from 猪圈设置",connt,3,3
ks=rs("投票开始")
jz=rs("投票结束")
dj=rs("等级")
rs.close
if aqjh_jhdj<dj then
 Response.Write "<script Language=JavaScript>alert('提示：等级不够，不能投票！');window.close();</script>"
 Response.End
end if
if cdate(ks)>now() then
 Response.Write "<script Language=JavaScript>alert('提示：投票没开始，不能颁奖！');window.close();</script>"
 Response.End
end if
if cdate(jz)<now() then
 Response.Write "<script Language=JavaScript>alert('提示：投票没结束，不能颁奖！');window.close();</script>"
 Response.End
end if
sql1="select * from 养猪风云榜 where 参赛人='"&name&"'"
rs.open sql1,connt,1,1
if rs.recordcount=0 then
 Response.Write "<script Language=JavaScript>alert('提示：此人未参与养猪大赛！');window.close();</script>"
 Response.End
else
 rs.close
 sql2="select * from 养猪投票者 where 投票ID="&aqjh_id
 rs.open sql2,connt,1,1
 if rs.recordcount<>0 then
   Response.Write "<script Language=JavaScript>alert('提示：不得重复投票！');window.close();</script>"
   Response.End
 else
   sql3="update 养猪风云榜 set 票数=票数+1 where 参赛人='"&name&"'"
   connt.execute(sql3)
   sql4="INSERT INTO 养猪投票者 (投票ID,姓名) VALUES ("&aqjh_id&",'"&aqjh_name&"')"
   connt.execute(sql4)
   Response.Write "<script Language=JavaScript>alert('提示：谢谢你投我的票！');window.close();</script>"
   Response.End
 end if
end if
%>