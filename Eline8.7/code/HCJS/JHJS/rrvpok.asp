<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
'http = Request.ServerVariables("HTTP_REFERER") 
'if InStr(http,"hcjs/jhjs")=0 then 
'	Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
'	Response.End 
'end if
vhtype=cint(trim(Request.Form("type")))
vhid=trim(Request.Form("id"))
if not isnumeric(vhid) then
	Response.Write "<script language=javascript>{alert('滚吧，你想做什么？你想捣乱呀！！！');window.close();}</script>" 
	Response.End 
end if
if vhtype=2 then
	tt=clng(trim(Request.Form("tt")))
	vhrnd=clng(trim(Request.Form("vhrnd")))/100
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT a,k FROM vh where id="&vhid,conn
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你想做什么？想捣乱吗？！');window.close();}</script>"
	Response.End
end if
if rs("k")>4 then
	conn.execute "delete * from vh where id="&vhid
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的座驾该报废了，被官府强制报废！');parent.location='vehicle.asp';}</script>"
	Response.End
end if
vhname=rs("a")
rs.close
rs.open "select h,j,m from b where a='"&vhname&"'",conn
randomize()
if vhtype=1 then
	vhh=int(rs("h")*0.3)
	vhm=int(rs("m")*0.6+0.5)
	vhj=int(rs("j")*(rnd()*0.1+0.9))
else
	vhtt=int(rs("m")/vhrnd+sqr(rs("h"))*vhrnd)
	if vhtt<>tt then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>{alert('滚吧，你想做什么？你想作弊呀！！！');window.close();}</script>" 
		Response.End 
end if
	vhh=int(rs("h")*vhrnd+0.5)
	vhm=int(rs("m")*vhrnd+0.5)
	vhj=int(rs("j")*(rnd()*0.24+2*vhrnd))
if rnd()<0.2 then vhj=0
end if
rs.close
rs.open "select 银两,金币 from 用户 where 姓名='"&sjjh_name&"'",conn
conn.execute "update 用户 set 银两=银两+"&vhh&",金币=金币+"&vhm&" where 姓名='"&sjjh_name&"'"
conn.execute "delete * from vh where id="&vhid
rs.close
set rs=nothing
conn.close
set conn=nothing
if vhj=0 then
	Response.Write "<script Language=Javascript>{alert('英雄没钱,无地之容,为了面子你把爱车处理,掉了得到银两"&vhh&"两，金币"&vhm&"个!');parent.location='vehicle.asp';}</script>"
else
	Response.Write "<script Language=Javascript>{alert('英雄没钱,无地之容,为了面子你把爱车处理,掉了得到银两"&vhh&"两，金币"&vhm&"个!');parent.location='vehicle.asp';}</script>"

end if
%>






