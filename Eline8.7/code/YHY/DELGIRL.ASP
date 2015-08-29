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
user=Session("sjjh_name")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("sjjh_usermdb")
conn.open connstr
sql="select * from 用户 where 姓名='" & user & "'"
set rs=conn.execute(sql)
sex=rs("性别")
if sex<>"女" then 
response.write "你有没有搞错呀，怡红院里可不要男的哦。<a href=index.asp>返回妓院</a>"
response.write "<br><a href=index.asp>返回</a>"
response.end
end if

%>
<!--#include file="jiu.asp"-->
<%
sql="select * from 名妓 where 姓名='" &user& "'"
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
response.write "你有没有搞错呀，本院哪有这个姑娘呀，你到别家去赎身吧<a href=index.asp>返回妓院</a>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("sjjh_usermdb")
conn.open connstr
sql="select * from 用户 where 姓名='" & user & "'"
set rs=conn.execute(sql)
yl=rs("银两")
if yl<5000000 then
response.write "你的腰包没有500万两白花银，就想来赎身呀，门都没有!<a href=index.asp>返回妓院</a>"
	response.end

else
sql="update 用户 set 银两=银两- 5000000   where 姓名= '"&user&"'"
set rs=conn.execute(sql)
connt.execute("delete * from 名妓 where 姓名='"&user& "'") %>
<!--#include file="connpic.asp"-->


<%  conn.execute("delete * from pic where 姓名='"&user& "'") 

	response.write "这是你的卖身契，你现在自由了，欢迎你在没钱的时候在来我们这工作！<a href=index.asp>返回妓院</a>"  
	response.end
conn.close
connt.close
set rs=nothing
end if
%>