<%
Sub Msg (v)
 Response.Write "<script Language=JavaScript>alert('" & v & "');history.back();</script>"
 Response.End
End Sub
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if aqjh_grade<10 or aqjh_name<>application("aqjh_user") then
	Response.Write "<b>[操作失败]</b><p>你没有更改投票时间的权限！"
	Response.End
end if
tp_dj=Trim(Request.Form("tp_dj"))
tp_start=Trim(Request.Form("tp_start"))
tp_end=Trim(Request.Form("tp_end"))
bm_start=Trim(Request.Form("bm_start"))
bm_end=Trim(Request.Form("bm_end"))
if IsNumeric(tp_dj)=false then
	Msg "[操作失败]投票所需要等级格式错误！"
end if
if IsDate(bm_start)=false or len(bm_start)<>19 then
	Msg "[操作失败]开始报名时间格式错误"
end if
if IsDate(tp_start)=false or len(tp_start)<>19 then
	Msg "[操作失败]开始投票时间格式错误！"
	Response.end
end if
if IsDate(bm_end)=false or len(bm_end)<>19 then
	Msg "[操作失败]终止报名时间格式错误！"
end if
if IsDate(tp_end)=false or len(tp_end)<>19 then
	Msg "[操作失败]终止投票时间格式错误！"
end if
if CDate(tp_end)<=CDate(tp_start) then
	Msg "[操作失败]终止投票时间必须晚于开始投票时间"
end if
if CDate(bm_end)<=CDate(bm_start) then
	Msg "[操作失败]终止报名时间必须晚于开始报名时间"
end if
if CDate(tp_start)<=CDate(bm_start) then
	Msg "[操作失败]开始投票时间必须晚于开始报名时间"
end if
if CDate(tp_end)<=CDate(bm_end) then
	Msg "[操作失败]终止投票时间必须晚于终止报名时间"
end if
%>
<!--#include file="data.asp"-->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
sql="update 猪圈设置 set 投票开始='" & tp_start & "',投票结束='" & tp_end & "',报名开始='"&bm_start&"',报名结束='"&bm_end&"',等级="&tp_dj
set rs=connt.execute(sql)
set rs=nothing
connt.close
set connt=nothing
Response.Write "<script Language=JavaScript>alert('设置成功');window.location.href ='hfyb.asp';</script>"
Response.End
%>