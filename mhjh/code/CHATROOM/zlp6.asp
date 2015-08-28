<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=Session("yx8_mhjh_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
if Application("yx8_mhjh_klt4")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你动作或者反应太慢，这种战利品已经被人拿走了！');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
	Response.end
end if
conn.execute "update 用户 set 体力=体力-"&Application("yx8_mhjh_klt4")&" where 姓名='"&username&"'"
rst.open "select 体力,门派 FROM 用户 WHERE 姓名='"&username&"'",conn
if rst("门派")="十八地狱" then
	Response.Write "<script Language=Javascript>alert('提示：你小子是不是没事找事呀，鬼魂还想得便宜啊！');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
	Application("yx8_mhjh_klt4")=0
	Response.end
end if
if rst("体力")<0 then
	conn.execute "update 用户 set 状态='死亡',被杀=被杀+1 where 姓名='"&username&"'"
conn.execute "insert into 英烈堂(时间,死者,凶手,死因) values(now(),'"&username&"','被怪物打死','体力不足累死')"
	kl="<img src='data/kl.gif'>太不幸了，<font color=0000FF>"&username&"</font>抢夺战利品时遇到强敌,由于体力不足而死,有够惨的!"
	Response.Write "<script Language=Javascript>alert('提示：你贪财不要命，体力太弱，丢了小命，去复活吧！没能耐别乱闯！');</script>"
	session.Abandon
	response.write "<script language=javascript>parent.top.location.href='chaterror.asp?id=009';</script>"
else
rst.close 
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.cursorlocation=3 
rst.Open "select * from 商店",conn
randomize       
r=150
id=int((r-1)*rnd)+1 
rst.close
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 商店 where id="&id,conn
ming=rst("名称")
rst.Close
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 物品 where 所有者='"&username&"' and 名称='"&ming&"'",conn
if rst.EOF or rst.BOF then
conn.Execute "insert into 物品(名称,属性,体力,内力,攻击,防御,特效,数量,价格,所有者,寄售,装备) values('"&rst("名称")&"','"&rst("属性")&"',"&rst("体力")&","&rst("内力")&","&rst("攻击")&","&rst("防御")&",'"&rst("特效")&"',1,"&rst("价格")\2&",'"&username&"',false,false)"
else
         conn.Execute "update 物品 set 数量=数量+1 where 所有者='"&username&"' and 名称='"&ming&"'"
	kl="<font color=0000FF>"&username&"</font>动作迅速，成功抢得武术家掉下的<img src=IMAGE/IMAGE/wg7.gif>" & ming &"！"
	Application("yx8_mhjh_klt4")=0
rst.Close
set rst=nothing
conn.Close
set conn=nothing
end if
end if
says="<font color=red><b>【抢战利品】</b></font>"&kl	
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
		newtalkarr(594)=name
		newtalkarr(595)="大家" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="008000" 
		newtalkarr(599)=""&says&""
        newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr
%>
<script Language=Javascript>alert('提示：抢夺战利品成功,注意看聊天室内的江湖通报！');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>
