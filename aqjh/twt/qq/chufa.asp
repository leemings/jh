<%@ LANGUAGE=VBScript%>
<!--#include file="../../showchat.asp"-->
<%Response.Expires=0
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if aqjh_name<>Application("aqjh_user") then Response.Redirect "qqlist.asp"
userid=request("uid")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 内力,银两,武功 from 用户 where 姓名='"&userid&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=454"
	response.end
end if
if request("id")=1 then
conn.execute "update 用户 set 内力=内力-50000,银两=银两-500000,武功=武功-5000 where 姓名='"&userid&"'"
else
conn.execute "update 用户 set 内力=内力-2000000,银两=银两-20000000,武功=武功-20000 where 姓名='"&userid&"'"
end if
rs.close
'写入QQ数据库
Set rs1=Server.CreateObject("ADODB.RecordSet")
sql="select * from QQ where 奖励=0 and oicq="&request("oicq")
rs1.open sql,conn,3,2
if rs1.bof or rs1.bof then response.redirect"qqlist.asp"
rs1("处罚")=1
if request("id")=1 then
rs1("处罚说明")="罚银两50万，内力5万，武功5千"
mess="罚银两50万，内力5万，武功5千"
else
rs1("处罚说明")="罚银两2000万，内力200万，武功2万"
mess="罚银两2000万，内力200万，武功2万"
end if
rs1.update
rs1.close
call showchat("<font color=red>【QQ宣传奖励】</font>"&userid&"宣传错误，被"&mess)
response.redirect"qqlist.asp?id=3"
%>