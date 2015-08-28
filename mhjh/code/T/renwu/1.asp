<!--#include file="../../config.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
timetmp=now()
my=session("yx8_mhjh_username")
if my="" then Response.Redirect "../error.asp?id=016"
if session("yx8_mhjh_usercorp")="十八地狱" then%>
<script language=vbscript>
MsgBox "错误！鬼魂禁止入内！"
location.href = "javascript:history.back()"
</script>
<%response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="select 任务时间,任务,风铃 from 用户 where 姓名='" & my & "'"
set rs=conn.execute(sql)
if rs("风铃")<1 then%>
<script language="vbscript">
  alert("你连一个风铃都没有,怎么来执行这个任务！")
location.href = "javascript:self.close()"
</script>
<%
else
if rs("风铃")>0 then
sql="update 用户 set 风铃=风铃-1,任务时间='" & timetmp & "',任务='山谷寻宝' where 姓名='" & my & "'"
conn.execute sql
dim newtalkarr(600) 
   talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=username 
		newtalkarr(595)="大家" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)="<font color=red>【江湖通报】</font><font color=blue>"&my&"使用了一个风铃,成功进入魔幻山谷,行动开始时间是" & timetmp & "让我们祝他好运吧!</font>" 
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
end if
end if
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>NPC</title>
<link rel="stylesheet" href="../../Style2.css">
</head>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20%"></td>
    <td width="46%"><img border="0" src="../../pic/OLDMAN.jpg"></td>
    <td width="33%"><font color="#800000">我是少林的</font><font color="#FF0000">伏魔练武师：</font><font color="#800000"><br>
      <br>
      <br>
      &nbsp;&nbsp;&nbsp; </font><a href="lian/1.asp"><font color="#800000">现在你可以在这练习打木桩了</font></a><font color="#800000"><br> 
      <br>
      <br>
      &nbsp;&nbsp;&nbsp; </font><a href="lian/d1.asp" target="_blank"><font color="#800000">如果能给我点好酒，就送你一样东西</font></a><font color="#800000"><br> 
      <br>
      &nbsp;&nbsp;&nbsp; <br>
      &nbsp;&nbsp;&nbsp; </font><a href="lian/d2.asp"><font color="#800000">看看，能不能从老和尚身上搜点东西<br>
      <br>
      <br>
      </font></a><font color="#800000"><a href="lian/d2.asp">&nbsp;&nbsp;&nbsp; </a></font><font color="#FF0000"><b>[每进行一次都要耗费一张驿卷]</b></font></td>
  </tr>
</table>
