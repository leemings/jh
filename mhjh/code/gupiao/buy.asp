<!--#include file="conn.asp"-->
<% 
if session("yx8_mhjh_username")="" then'1
  call endinfo("不能进入，你还没有登录")
end if

sid=cstr(abs(Request.Form ("sid")))
ushare=abs(Request.Form ("ushare"))
if ushare>10000000 then
  call endinfo("您一次最多可以买一千万支股票！")
  response.end
end if

sql= "select * from 股票 where sid="&sid        
set rs=conn.execute(sql) 
response.write rs("当前价格")
response.end

if rs("当前价格")<=1 OR (rs("当前价格")-rs("开盘价格"))/rs("开盘价格")>=0.15 then'2
  call endinfo("对于停牌、涨停板或者跌停板的股票是不能进行买操作的哦：） ")
else
  yx8_mhjh_username=session("yx8_mhjh_username")
  sql="select * from 客户 where 帐号='"&session("yx8_mhjh_username")&"'"
  set rs=conn.execute(sql)
  username=rs("帐号")
  nowmoney=rs("资金")

  set rs_s=conn.execute("select 流通股票,开盘价格,当前价格,企业 from 股票 where sid="&sid)
  set rs_u1=conn.execute("select 持股数,帐号,买入价格 from 大户 where sid="&sid)
  nowp=rs_s("当前价格")
  dsshare=rs_s("流通股票")-ushare
  tot=rs_s("流通股票")

  if dsshare<0 then
    call endinfo("没有足够的股票数量！")
    rs_s.close
    rs.close
    response.end
  elseif nowmoney<ushare*rs_s("当前价格")*1.01 then
    call endinfo("您的资金不足！")
    rs_s.close
    rs.close
    response.end
  end if

  sql = "select count(*) as num from 事件 "
  set rs=conn.execute(sql)
  num=rs("num")
  if num>=40 then '*
    Set rs = Server.CreateObject("ADODB.recordset")
    sql = "select top 1 * from 事件 order by 原因时间 asc"
    rs.open sql,conn,3,2
    rs.delete
  end if '*

  if nowmoney>=ushare*rs_s("当前价格")*1.01 and dsshare>=0 then'3
    delmoney=ushare*rs_s("当前价格")
    set rs_u=conn.execute ("select 持股数,买入价格 from 大户 where sid="&sid&" and 帐号='"&username&"'")
    if rs_u.eof then'4
      Randomize
      s=Rnd
      if s<=0.4 then
        s=s+0.5
      end if
      sprice=nowp*(1+ushare*s/(tot*5))
      sql="update 股票 set 当前价格="&sprice&", 流通股票="&dsshare&",交易量=交易量+"&ushare&" where sid="&sid
      conn.execute sql
      sql="insert into 大户 (帐号,sid,买入价格,平均价格,持股数,企业,时间) values ('"&username&"','"&sid&"','"&nowp&"','"&nowp&"','"&ushare&"','"&rs_s("企业")&"','"&date()&"')"
      conn.execute sql
      sql="update 客户 set 资金=资金-"&delmoney&"*1.01 where  帐号='"&username&"'"
      conn.execute sql
      sql="update 总资金 set 总资金=总资金+"&delmoney&"*0.99 where  帐号='"&username&"'"
      conn.execute sql  
      zd=ccur((formatcurrency(sprice,3)-formatcurrency(nowp,3))/formatcurrency(rs_s("当前价格"),3))
      mess="<font color=#990000>"&username&"买入"&rs_s("企业")&" "&ushare&" 股，现价上涨 "&formatpercent(zd,2,-1)&"</font>"
      sql="insert into 事件(原因,原因时间) values('"&mess&"','"&now()&"' )"
      conn.execute sql  
    else
      yuan=rs_u("持股数")
      Randomize
      s=Rnd

      inprice=rs_u1("买入价格")
      if s<=0.4 then
        s=s+0.5
      end if
      sprice=nowp*(1+ushare*s/(tot*5))
      uprice=(rs_s("当前价格")*ushare+rs_u("买入价格")*yuan)/(ushare+yuan)
      sql="update 股票 set 当前价格="&sprice&", 流通股票="&dsshare&",交易量=交易量+"&ushare&" where sid="&sid
      conn.execute sql
      sql="update 大户 set 平均价格="&uprice&",买入价格="&nowp&",持股数="&rs_u("持股数")&"+"&ushare&" where 帐号='"&username&"' and sid="&sid
      conn.execute sql
      sql="update 客户 set 资金=资金-"&delmoney&"*1.01 where  帐号='"&username&"'"
      conn.execute sql

      sql="update 总资金 set 总资金=总资金+"&delmoney&"*0.99 where  帐号='"&username&"'"
      conn.execute sql
      zd=ccur((formatcurrency(sprice,3)-formatcurrency(nowp,3))/formatcurrency(rs_s("当前价格"),3))
      mess="<font color=#990000>"&username&"再次买入"&rs_s("企业")&" "&ushare&" 股，现价上涨 "&formatpercent(zd,2,-1)&" </font>"
      sql="insert into 事件(原因,原因时间) values('"&mess&"','"&now()&"' )"
      conn.execute sql  
    end if'4
  end if'3
dim newtalkarr(600) 
   talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=username 
		newtalkarr(595)="大家" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)="<font color=red>【股票收购】</font><font color=blue>"&mess&"</font>" 
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<link rel="stylesheet" type="text/css" href="../style1.css">
<%
  call endinfo("股票购买成功！")
end if

%>
<%conn.close
set conn=nothing
sub endinfo(message) 
'-------------------------------信息提示-------------------------------
%>
<table width="<%=TableWidth%>" border=0 cellspacing=1 cellpadding=3 align=center bgcolor="<%=Tablebackcolor%>"><tr bgcolor="<%=Tabletitlecolor%>" ><td align=center height=26 bgcolor="<%=Tabletitlecolor%>"><b>信息提示</b></td></tr><tr><td align=center height=70 bgcolor="<%=aTabletitlecolor%>"><%=message%><br></td></tr><tr><td align=center height=26 bgcolor="<%=Tabletitlecolor%>"><a href="stock.asp">返回</a></td></tr></table>
<%end sub
%>

