<!--#include file="conn.asp"-->
<!--#include file="chkmaster.asp"-->
<% 
if   session("aqjh_name") ="" then
response.redirect "../login.asp"
else
 
if not master then
call endinfo("你没有资格操作")
else
  
sql = "select count(*) as num from 事件 "
set rs=conn.execute(sql)
num=rs("num")
if num>=40 then '*
Set rs = Server.CreateObject("ADODB.recordset")
sql = "select top 1 * from 事件 order by 原因时间 asc"
rs.open sql,conn,3,2
rs.delete
end if '*

tt=now()
sql2="select count(*) as toto from 股票"   
set rsst=conn.execute(sql2)  
toto=rsst("toto")
rsst.close
for sj=1 to toto
sql= "select * from 股票"
set rs=conn.execute(sql) 
if (rs("当前价格")-rs("开盘价格"))/rs("开盘价格")>-0.15 then 
sql="update 股票 set 当前价格=当前价格*0.97 where sid="&sj
conn.execute sql 
end if
rs.movenext
next 
mess="经济发展过热，通货膨胀严重，股票市场所有股票价格下挫 3 %"
sql="insert into 事件(原因,原因时间) values('"&mess&"','"&tt&"' )"
conn.execute sql  



	sj=formatdatetime(time(),3)
	sj="<font color=#FF00FF size=1>(" & sj & ")</font>"

CloseDatabase	
Response.Redirect "stock.asp"

end if
end if

 sub endinfo(message) 
'-------------------------------信息提示-------------------------------
%>
<table width="<%=TableWidth%>" border=0 cellspacing=1 cellpadding=3 align=center bgcolor="<%=Tablebackcolor%>"><tr bgcolor="<%=Tabletitlecolor%>" ><td align=center height=26 bgcolor="<%=Tabletitlecolor%>"><b>信息提示</b></td></tr><tr><td align=center height=70 bgcolor="<%=aTabletitlecolor%>"><%=message%><br></td></tr><tr><td align=center height=26 bgcolor="<%=Tabletitlecolor%>"><a href="stock.asp">返回</a></td></tr></table>
<%end sub
%>