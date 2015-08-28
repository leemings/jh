<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
lastip=Request.ServerVariables("QUERY_STRING")
if instr(lastip,"or") then response.end
my=session("yx8_mhjh_username")
if my="" then response.redirect "../error.asp?id=016"
id=request("id")
for each element in Request.Form
elevalue=Request.Form(element)
if instr(elevalue,"'")<>0 or instr(elevalue,chr(34))<>0 or instr(elevalue,"\")<>0 or instr(elevalue,";") or instr(elevalue,"?") or instr(elevalue,"or") or instr(elevalue,"&") or instr(elevalue,"#")then Response.Redirect "../error.asp?id=056"
next
%>
<!--#include file="data.asp"-->
<%
sql="select 配偶 from 用户 where 姓名='" & my & "'"
set rs=conn.execute(sql)
po=rs("配偶")
if po="无" or po=""  then
%>
<script language=vbscript>
 MsgBox "你还没有老婆，不能买房屋!"
 location.href = "../11.htm"
</script>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
response.end
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<!--#include file="data1.asp"-->
<%
sql="select * from 房屋 where 户主='" & my & "' or 伴侣='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
sql="select * from 房屋 where ID=" & id
rs=conn.execute(sql)
yin=rs("售价")
huzhu=rs("户主")
set rs=nothing
conn.close
set conn=nothing
%>
<!--#include file="data.asp"-->
<%
sql="select 银两 from 用户 where 姓名='" & my & "'"
set rs=conn.execute(sql)
if rs("银两")>=yin and my<>huzhu and huzhu="无" then
      sql="update 用户 set 银两=银两-" & yin & " where 姓名='" & my & "'"
      conn.execute(sql)
      set rs=nothing
      conn.close
      set conn=nothing
%>
<!--#include file="data1.asp"-->
<%
      sql="update 房屋 set 户主='" & my & "',伴侣='" & po & "' where ID=" & id
      conn.execute(sql)
      set rs=nothing
      conn.close
      set conn=nothing
Response.Write "<head><title>消息</title><link rel=stylesheet href='../chatroom/css.css'><script lanjuage=javascript>setTimeout('history.back();',3000);</script></script></head><body  bgcolor='"&bgcolor&"' background='"&bgimage&"' oncontextmenu='self.event.returnValue=false;' topmargin=100><p align=center>恭喜，你已经成功购买房屋！<br>三秒钟后自动返回<br><a href='javascript:history.back();'>返回</a></p></body>"
else
%>
<script language=vbscript>
 MsgBox "购买不成功，原因：你的银两不够。"
 location.href = "../11.htm"
</script>
<%
conn.close
response.end
end if
%>
<%else%> 
<script language=vbscript>
     MsgBox "您或您的伴侣已经买过房屋了。"
     location.href = "../11.htm"
</script>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
end if
%>