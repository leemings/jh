<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if Session("sjjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "请进入聊天室再进入股票！"
window.close()
</script>
<%Response.End
end if
if sjjh_name=""  then
  Response.Redirect "../error.asp?id=440"
elseif sjjh_grade<10 then
	Response.Redirect "../error.asp?id=439"
else%>
<!--#include file="jhb.asp"-->
<%
sql= "select * from 股票 where sid=1"        
set rss=conn.execute(sql)
if rss("状态")="开" then 
%>
<script language="vbscript">
  alert "目前股市已经是开盘状态了"
  location.href = "index.asp"
</script>
<%
elseif rss("日期")=date() then
%>
<script language="vbscript">
  alert "今天股市已经封盘不能再开了"
  location.href = "index.asp"
</script>
<%
else
for i=1 to 20  
sql= "select * from 股票 where sid="&i        
set rs=conn.execute(sql)   
sql="update 股票 set 开盘价格='"&rs("当前价格")&"',状态='开',日期='"&date()&"'where sid="&i
conn.execute sql
next
end if
end if
conn.close
Response.Redirect "index.asp"

%>

