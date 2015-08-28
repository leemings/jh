<%
Response.Buffer=true
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "error.asp?id=016"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="SELECT 性别,道德 FROM 用户 WHERE 姓名='"&username&"'"
set rs=conn.execute(sql)
sex=rs("性别")
meimao=rs("道德")
if sex<>"女" then 
response.write "你有没有搞错呀，我们女性画廊里可不要男的哦。"
response.write "<br><a href=index.asp>返回</a>"
response.end
end if
if meimao<1000 then 
response.write "你有没有搞错呀，你长的这么丑还来这儿，你的道德要大于1000才可以的。"
response.write "<br><a href=index.asp>返回</a>"
response.end
conn.colse
set rs=nothing
end if%>
<!--#include file="jiu.asp"-->
<% 
if meimao>50000 then
   meimao=50000
end if
sql="select * from 名妓 where 姓名='" & username & "'"
set rs=connt.execute(sql)
if rs.bof or rs.eof then
sql="insert into 名妓(姓名,美貌度) values ('" & username & "'," & meimao & ")"
connt.execute sql
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="update 用户 set 银两=银两+500000 where 姓名='"& username &"'"
set rs=conn.execute(sql)
response.write "恭喜，你正式成为本院的姑娘，等到卖身银两50万！"
response.end
response.write "<br><a href=index.asp>返回</a>"
else
response.write "你已经是本女性画廊的姑娘了，怎么还来登记呀！"
response.write "<br><a href=index.asp>返回</a>"
response.end
connt.close
conn.close
set rs=nothing
end if
%>
