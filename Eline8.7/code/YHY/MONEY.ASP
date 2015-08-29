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
response.write "你有没有搞错呀，男的也想来领工资呀？<a href=index.asp>返回妓院</a>"
response.write "<br><a href=index.asp>返回</a>"
response.end
end if


%>
<!--#include file="jiu.asp"-->
<%
sql="select * from 名妓 where 姓名='" &user& "'"
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
response.write "你有没有搞错呀，本院哪有这个姑娘呀，你到别家去领钱吧<a href=index.asp>返回妓院</a>"
	response.end
end if
mdate=rs("money")
mtime=rs("time")

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("sjjh_usermdb")
conn.open connstr
sql="select * from 用户 where 姓名='" & user & "'"
set rs=conn.execute(sql)

meimao=rs("魅力")

money=meimao*mtime/10
if day(mdate)<day(now()) or month(mdate)<month(now()) or year(mdate)<year(now()) or isnull(mdate) then
		conn.execute("update 用户 set 银两=银两+"&money&" where 姓名='"&user&"'"),null
       
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open "select * from 名妓 where 姓名='"&user&"'",connt,3,2

rs("money")=now()
rs.update

       %>
    你从妓院老板那儿领到了今天辛苦得来的<%=money %>两血汗钱，小心拿好啊。<A onclick=javascript:history.back(); 
href="#">返回</A><br>
    <%
else %>

你今天已经领过钱啦,好好工作吧,我不会亏待你得.<A onclick=javascript:history.back(); 
href="#">返回</A>
<%
end if
%><!--#include file="connpic.asp"--><%  
conn.close
connt.close
set rs=nothing

%>