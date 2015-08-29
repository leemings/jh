<!--#include file="conn.asp"-->
<!--#include file="z_fy_Conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_fy_const.asp"-->
<%
'=========================================================
' Plusname: 法院监狱第三版
' File: z_fy_looktj.asp
' For Version: 6.0sp2
' Date: 2003-2-10
' Script Written by Hero20000
'=========================================================
dim fanren,fysql,fyrs
fanren=checkstr(request("name"))
call getfyconfig()
%>
<html>
<head>
<title>查看<%=fanren%>留言（事件）</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body background="Images/img_fy/Bg001.gif" >
<!-- #include file="inc/FORUM_CSS.asp" -->    
<table border=0  cellspacing=1 cellpadding=3 align="center" width=60% class=tableborder1>
<%
select case  checkstr(request("action"))
case "list" 
  call listh()
case "delmy"
  call delmy()
case "delall"
  call delall()
case else
if log_set(0)="1" then
fysql="delete from log where class='监狱' and tousername='"&fanren&"' and title<>'入狱' and id not in (select top "&clng(log_set(1))&" id from log where class='监狱' and tousername='"&fanren&"' and title<>'入狱' order by id desc)"
connfy.execute(fysql)
end if
  fysql="SELECT username,title,content,dateandtime FROM log WHERE class='监狱' and tousername='"&fanren&"' order by dateandtime desc"
  set fyrs=connfy.execute(fysql)
  if fyrs.eof then
    response.write "<tr><td class=tablebody1 align=center>呵呵，监狱的事件/留言中没有您要的 "&fanren&" 记录哦！</td></tr>"
  else
    do while Not fyRS.Eof 
     response.write "<tr><td  height=24 width=40% class=tablebody2>日期："&fyrs(3)&"</td>"
     response.write "<td height=24 class=tablebody2 >事件："&fyrs(1)&"</td>"
     response.write "<td height=24 width=40% class=tablebody2 >"
         select case fyrs(1)
           case "入狱"
        response.write "管理员/法官："
      case "出狱"
        response.write "管理员/法官："
      case "保释"
        response.write "恩人："
      case "抢劫"
        response.write "强盗："
      case "探监"
        response.write "留言人："
    end select 
    if fyrs(1)="抢劫" and qj_set(3)<>"1" then
      response.write " 蒙面大盗</td></tr>"
    else  
    response.write " "&fyrs(0)&"</td></tr>"
    end if
     response.write "<tr><td class=tablebody1 colspan=3 >详情：<br>"&fyrs(2)&"<br></td></tr>"
    fyRS.MoveNext
    Loop
  end if
end select

%>
</table>
 
<br>
<div align="center">
<%
if membername=fanren or master or jymaster then 
response.write "【<a href=z_fy_looktj.asp?action=delmy&name="&fanren&">"
 if master or jymaster then 
  response.write "清空该用户监狱记录</a>】<br>"
 else
   response.write "清空我的记载</a>】<br>"
 end if
end if
if master then response.write  "【<a href=z_fy_looktj.asp?action=delall>管理员清空所有记载</a>】<br>"
%><input type="button" value="关闭" onClick="window.close()" name="button">
</div>
<% 
sub delmy()   
 if membername<>checkstr(request("name")) and  (not master) and (not jymaster) then   
  response.write "<tr><td class=tablebody1>对不起，你删别人的记录干什么？</td></tr>"   
 else
  if membername=checkstr(request("name")) then   
     fysql="delete * from log where class='监狱' and title<>'入狱' and tousername='"&membername&"'"
  else
     fysql="delete * from log where class='监狱' and tousername='"&checkstr(request("name"))&"'"
  end if
  connfy.execute(fysql)   
  Response.Redirect "z_fy_looktj.asp"   
 end if   
end sub   
   
sub delall()   
 if not master then   
   response.write "<tr><td class=tablebody1>对不起，你不是管理员呀！</td></tr>"   
 else   
   fysql="delete * from log where class='监狱' "   
   connfy.execute(fysql)   
   Response.Redirect "z_fy_looktj.asp"   
 end if   
end sub   
   
sub listh()   
  fysql="SELECT DISTINCT tousername FROM log WHERE class='监狱' and (title='出狱' or title='保释')"   
  set fyrs=connfy.execute(fysql)   
  if fyrs.eof then   
    response.write "<tr><td class=tablebody1 align=center>呵呵，没有记录哦！</td></tr>"   
  else   
    dim unum   
    unum=0   
    response.write "<tr>"   
    do while Not fyRS.Eof  
    unum=unum+1   
     response.write "<td class=tablebody1 align=center>【<a href=z_fy_looktj.asp?name="&fyrs(0)&" title='点击查看留言'>"&fyrs(0)&"</a>】</td>"   
    if (unum mod 4=0) then response.write "</tr><tr>"   
    fyRS.MoveNext   
    Loop   
    response.write "</tr>"   
  end if   
end sub   
  
 %>   
</body></html>