<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../showchat.asp"-->
<!--#include file="../const2.asp"-->
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
dd=day(now())
jl_name=trim(request("jl_name"))
if dd<>jl_dd then 
	Response.Write "<script Language=Javascript>alert('提示：现在还不是"&jl_dd&"号！！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
%>
<html>
<head>
<title><%=Application("aqjh_chatroomname")%>处理本月奖励</title>
<LINK href="css/css.css" type=text/css rel=stylesheet>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "Select * from 用户 Where 姓名='"&jl_name&"'",conn
if rs.eof then
  	Response.Write "<script Language=Javascript>alert('提示：找不到用户！！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
else
'处理月积分奖励
rs.close
rs.open "Select top "&jl_yjf_top&" * from 用户 where 门派<>'官府' order by mvalue desc",conn,1,3
top=1
do while not rs.eof
   if rs("姓名")=jl_name and top<=jl_yjf_top then
        top1=jl_yjf_top-top+1
        conn.Execute("update 用户 set 会员金卡=会员金卡+"&top1*jl_yjf_jk&" where 姓名='"&jl_name&"'")
        call showchat("<img src=../chat/img/jk.gif><font color=ff0000>【本月月积分奖励】</font>"&jl_name&"为本月月积分排行第["&top&"]名,今奖励会员金卡["&top1*jl_yjf_jk&"]元。")
   end if
   top=top+1
   rs.movenext
loop
'处理拉人排行
rs.close
sql="SELECT 介绍人,count(id) as num FROM 用户 where allvalue>"&jl_allvalue&" and 介绍人='"&jl_name&"' group by 介绍人 order by count(id) desc "
rs.open sql,conn,1,1
if not rs.eof then
   num=rs("num")
   conn.Execute("update 用户 set 会员金卡=会员金卡+"&num*jl_lr_jk&",银两=银两+"&num*jl_lr_yl&" where 姓名='"&jl_name&"'")
   call showchat("<img src=../chat/img/menoy.gif><font color=ff0000>【本月拉人奖励】</font>"&jl_name&"为江湖一共拉了["&num&"]人(达到条件)，江湖特奖励会员金卡["&num*jl_lr_jk&"]元,白银["&num*jl_lr_yl&"]两。")
end if
  	Response.Write "<script Language=Javascript>alert('提示：本月奖励["&jl_name&"]用户处理完毕！！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
%>
</body></html>