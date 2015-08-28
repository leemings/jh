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
<!--#include file="data1.asp"-->
<%
sql="select * from 房屋 where ID=" & id
set rs=conn.execute(sql)
yin=rs("售价")
huzhu=rs("户主")
blv=rs("伴侣")
zt=rs("状态")
if rs.eof and rs.bof then
%>
<script language=vbscript>
 MsgBox "房子不存在"
 location.href = "../11.htm"
</script>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
response.end
end if
if rs("银两")<0 or rs("伴侣银两")<0 or rs("数量")<0  then
%>
<script language=vbscript>
       MsgBox "错误：由于你在房屋里作弊，你的房屋被官府收回，你被官府的刽子手斩首！"
       location.href = "../exit.asp"
</script>
<%
sql="update 房屋 set 户主='无',伴侣='无',数量=0,银两=0,伴侣银两=0 where ID=" & id
conn.execute(sql)
sql="delete * from 物品 where 拥有者='" & huzhu & "' or 拥有者='" & blv & "'"
conn.execute(sql)
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<!--#include file="data.asp"-->
<%
sql="update 用户 set 状态='死亡',银两=-200000,体力=体力-100000,内力=内力-100000 where 姓名='" & my & "'"
conn.execute(sql)
sql="insert into 英烈堂(死者,时间,凶手,死因) values ('" & my & "',now(),'刽子手','在房间作弊')"
conn.execute(sql)
conn.close
set conn=nothing
response.end
end if
%>
<!--#include file="data.asp"-->
<%
if my=huzhu  and zt="正常" then
    sql="update 用户 set 银两=银两+" & yin & " where 姓名='" & my & "'"
    conn.execute(sql)
    conn.close
    set conn=nothing
%>
<!--#include file="data1.asp"-->
<%
       sql="update 房屋 set 户主='无',伴侣='无',数量=0,银两=0,伴侣银两=0 where ID=" & id
       conn.execute(sql)
       conn.close
       set conn=nothing
Response.Write "<head><title>消息</title><link rel=stylesheet href='../chatroom/css.css'><script lanjuage=javascript>setTimeout('history.back();',3000);</script></script></head><body  bgcolor='"&bgcolor&"' background='"&bgimage&"' oncontextmenu='self.event.returnValue=false;' topmargin=100><p align=center>恭喜，你已经成功出售房屋！<br>三秒钟后自动返回<br><a href='javascript:history.back();'>返回</a></p></body>"
else
%>
<script language=vbscript>
 MsgBox "交易不成功，原因：你不是这个房子的主人或是你的房子出了点状况!"
 location.href = "../11.htm"
</script>
<%
end if
%>