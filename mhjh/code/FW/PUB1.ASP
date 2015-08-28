<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
lastip=Request.ServerVariables("QUERY_STRING")
if instr(lastip,"or") then response.end
my=session("yx8_mhjh_username")
if my="" then Response.Redirect "../error.asp?id=016"
%>
<!--#include file="data.asp"-->
<%
date1=request.form("date")+0
time1=request.form("time")+0
for each element in Request.Form
elevalue=Request.Form(element)
if instr(elevalue,"'")<>0 or instr(elevalue,chr(34))<>0 or instr(elevalue,"\")<>0 or instr(elevalue,";") or instr(elevalue,"?") or instr(elevalue,"or") or instr(elevalue,"&") or instr(elevalue,"#")then Response.Redirect "../error.asp?id=056"
next
if date1=0 and time1=0 then
%>
<script language=vbscript>
Msgbox "请输入正确的数字!!"
history.back(-1)
</script>
<%
else
	yin=(date1*24+time1)*50
		shi=date1+0.0416*time1
	sql="select * from 用户 where 姓名='" & my & "' and  状态<>'逮捕' and 状态<>'死亡' and 状态<>'入狱'"
	set rs=conn.execute(sql)
	if rs.eof or rs.bof then
		mess=my & "，你不能操作！"
		else
			sql="update 用户 set 最后登录时间=now()+"&shi&",状态='眠',体力=体力+"&yin&",内力=内力+"&yin&" where 姓名='" & my & "'"
			conn.execute sql
			mess=my & "，你已经成功的在家里休息了，请在" & date1 & "天" &time1& "小时后使用该帐号！"
       Response.Write "<head><title>消息</title><link rel=stylesheet href='../chatroom/css.css'><script lanjuage=javascript>setTimeout('history.back();',3000);</script></script></head><body  oncontextmenu='self.event.returnValue=false;' topmargin=100><p align=center><br>"&mess&"<br><br><a href='javascript:history.back();'>返回</a></p></body>" 
   Session.Abandon
%>

<%
end if
rs.close
set rs=nothing
end if
conn.close
set conn=nothing
%>