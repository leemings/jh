<%@ LANGUAGE=VBScript codepage ="936" %>
<% 
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
jiutian=Session("jiutian")
nickname=Session("sjjh_name")
dj=Session("sjjh_jhdj")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("sjjh_usermdb")
conn.open connstr
sql="select * from 用户 where 姓名='" & nickname & "'"
set rs=conn.execute(sql)
sex=rs("性别")
if sex<>"女" then 
response.write "你有没有搞错呀，怡红院里可不要男的哦。"
response.write "<br><a href=index.asp>返回</a>"
response.end
end if
if sjjh_grade<2 then 
response.write "江湖小辈，本怡红院不要没名望的姑娘！！！"
response.write "<br><a href=index.asp>返回</a>"
response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("sjjh_usermdb")
conn.open connstr
sql="SELECT * FROM 用户 WHERE 姓名='"&nickname&"'"
set rs=conn.execute(sql)
sex=rs("性别")
meimao=rs("魅力")
if meimao<1000 then 
response.write "你有没有搞错呀，你长的这么丑还来这儿。"
response.write "<br><a href=index.asp>返回</a>"
response.end
conn.close
set conn=nothing
rs.close
set rs=nothing
end if%>
<!--#include file="jiu.asp"-->
<% 
sql="select * from 名妓 where 姓名='" & username & "'"
set rs=connt.execute(sql)
if rs.bof or rs.eof then
sql="insert into 名妓(姓名,美貌度) values ('" & username & "'," & meimao & ")"
			connt.execute sql
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("sjjh_usermdb")
conn.open connstr
sql="SELECT * FROM 用户 WHERE 姓名='"&username&"'"
set rs=conn.execute(sql)
sql="update 用户 set 银两=银两+1500000 where 姓名='"& username &"'"
set rs=conn.execute(sql)
response.write "恭喜，你正式成为本院的姑娘，等到卖身银两1500000万！"

response.write "<br><a href=index.asp>返回</a>"
response.end
connt.close
set connt=nothing
conn.close
set conn=nothing
rs.close
set rs=nothing
else
response.write "你已经是本怡红院的姑娘了，怎么还来登记呀！"
response.write "<br><a href=index.asp>返回</a>"
response.end
connt.close
set connt=nothing
conn.close
set conn=nothing
rs.close
set rs=nothing
end if
%>