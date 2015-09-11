<!--#include file="conn.asp"-->
<% 

if session("aqjh_name")="" then
response.redirect "../login.asp"
response.end
else
%>
 
<%sid=Request.Form ("sid")
ushare=abs(Request.Form ("ushare2"))
sql= "select * from 股票 where sid="&sid        
set rs=conn.execute(sql) 
if rs("当前价格")<=1  or (rs("当前价格")-rs("开盘价格"))/rs("开盘价格")<=-0.15 then
call endinfo("对于停牌、涨停板或者跌停板的股票是不能进行卖操作的哦：）")

else
set rs=nothing
 
session("aqjh_name")=session("aqjh_name")
sql="select * from 客户 where 帐号='"&session("aqjh_name")&"'"
set rs=conn.execute(sql)
username=rs("帐号")
nowmoney=rs("资金")

set rs_s=conn.execute ("select 开盘价格,流通股票,当前价格,企业 from 股票 where sid="&sid)
set rs_u=conn.execute ("select 持股数,买入价格,平均价格 from 大户 where sid="&sid&" and  帐号='"&username&"'")

dsshare=rs_s("流通股票")+ushare
nowp=rs_s("当前价格")

addmoney=ushare*rs_s("当前价格")
nowp=rs_s("当前价格")
tot=rs_s("流通股票")

sql = "select count(*) as num from 事件 "
set rs=conn.execute(sql)
num=rs("num")
if num>=40 then '*
Set rs = Server.CreateObject("ADODB.recordset")
sql = "select top 1 * from 事件 order by 原因时间 asc"
rs.open sql,conn,3,2
rs.delete
end if '*

if rs_u.eof then 
call endinfo("您没有这种股票！")
else
dushare=rs_u("持股数")-ushare
if dushare<0 then
call endinfo("您的股票数不足！")
response.end

elseif dushare=0 then
Randomize
s=Rnd


sprice=nowp*(1-ushare/(tot*10))


shou=(rs_s("当前价格")-rs_u("平均价格"))*ushare
conn.execute "update 股票 set 日期="&date()&",交易量=交易量+"&ushare&",当前价格="&sprice&", 流通股票="&dsshare&" where sid="&sid
conn.execute "delete from 大户 where 帐号='"&username&"' and sid="&sid
sql="update 客户 set 资金=资金+"&addmoney&"*0.99 where 帐号='"&username&"'"
conn.execute sql
Set rs2 = Server.CreateObject("ADODB.recordset")
sql2="select 总资金 from 总资金 where 帐号='"&username&"'"
rs2.open sql2,conn
if (rs2("总资金")-addmoney*1.01)<=0 then 
sql="update 总资金 set 总资金=0 where  帐号='"&username&"'"
conn.execute sql
else
sql="update 总资金 set 总资金=总资金-"&addmoney&"*1.01 where  帐号='"&username&"'"
conn.execute sql
end if
rs2.close
set rs2=nothing

zd=ccur((formatcurrency(nowp,3)-formatcurrency(sprice,3))/formatcurrency(rs_s("当前价格"),3))
mess="<font color=#006600>"&username&"卖出"&rs_s("企业")&" "&ushare&" 股，现价下滑 "&formatpercent(zd,2,-1)&"</font>"

sql="insert into 事件(原因,原因时间) values('"&mess&"','"&now()&"' )"
conn.execute sql  
else
Randomize
s=Rnd

sprice=nowp*(1-ushare/(tot*10))

shou=(rs_s("当前价格")-rs_u("平均价格"))*ushare
conn.execute "update 股票 set 日期="&date()&",交易量=交易量+"&ushare&",当前价格="&sprice&", 流通股票="&dsshare&" where sid="&sid
conn.execute "update 大户 set 持股数="&dushare&" where 帐号='"&username&"' and sid="&sid
sql="update 客户 set 资金=资金+"&addmoney&"*0.99 where 帐号='"&username&"'"
conn.execute sql

sql2="select 总资金 from 总资金 where 帐号='"&username&"'"
Set rs2 = Server.CreateObject("ADODB.recordset") 
rs2.open sql2,conn
if (rs2("总资金")-addmoney*1.01)<=0 then 
sql="update 总资金 set 总资金=0 where  帐号='"&username&"'"
conn.execute sql
else
sql="update 总资金 set 总资金=总资金-"&addmoney&"*1.01 where  帐号='"&username&"'"
conn.execute sql
end if
rs2.close
set rs2=nothing

zd=ccur((formatcurrency(nowp,3)-formatcurrency(sprice,3))/formatcurrency(rs_s("当前价格"),3))
mess="<font color=#006600>"&username&"卖出"&rs_s("企业")&" "&ushare&" 股，现价下滑 "&formatpercent(zd,2,-1)&" </font>"

sql="insert into 事件(原因,原因时间) values('"&mess&"','"&now()&"' )"
conn.execute sql  

end if
call endinfo("股票卖出交易已成功提交！确定后返回")
end if
end if
end if

 sub endinfo(message) 
'-------------------------------信息提示-------------------------------
%>
<table width="<%=TableWidth%>" border=0 cellspacing=1 cellpadding=3 align=center bgcolor="<%=Tablebackcolor%>"><tr bgcolor="<%=Tabletitlecolor%>" ><td align=center height=26 bgcolor="<%=Tabletitlecolor%>"><b>信息提示</b></td></tr><tr><td align=center height=70 bgcolor="<%=aTabletitlecolor%>"><%=message%><br></td></tr><tr><td align=center height=26 bgcolor="<%=Tabletitlecolor%>"><a href="stock.asp">返回</a></td></tr></table>
<%end sub
%>
<%conn.close
set conn=nothing

%>