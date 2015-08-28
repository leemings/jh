<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
jstl=Application("yx8_mhjh_klt")
un=Session("yx8_mhjh_username")
if Application("yx8_mhjh_klt")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你动作或者反应太慢，机会已经错过了！')</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
	Response.end
end if
tempjs=int(abs(clng(Application("yx8_mhjh_klt"))))
'if tempjs<>jstl then
	'Response.Write "<script Language=Javascript>alert('提示：滚，你想搞什么，按回车键不行了!');</script>"
	'Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
	'response.end
'end if
randomize          
r=rst.recordcount 
id=int((r-1)*rnd)+1 
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
rst.open "select * from 商店" where id="&id
rst.Open sqlstr,conn
if rst.EOF or rst.BOF then
says="<font color=red><b>你运气不好,什么也没得到</b></font>"
else
commodityname=rst("名称")
commoditytype=rst("属性")
commodityhealth=rst("体力")
magic=rst("内力")
attack=rst("攻击")
defence=rst("防御")
especial=rst("特效")
commodityprice=rst("价格")
rst.Close
rst.Open "select 银两 from 用户 where 姓名='"&username&"' and 银两>10000,conn
if rst.EOF or rst.BOF then
msg="错误，你的钱好象没有带够哟！"
else
rst.Close
sqlstr="select * from 物品 where 名称='"&commodityname&"' and 所有者='"&un&"'"
rst.Open sqlstr,conn
if rst.EOF or rst.BOF then
sqlstr="insert into 物品(名称,属性,体力,内力,攻击,防御,特效,价格,数量,所有者,寄售,装备) values('"&commodityname&"','"&commoditytype&"',"&commodityhealth&","&magic&","&attack&","&defence&",'"&especial&"',"&commodityprice/2&",1,'"&un&"',False,False)"
else
sqlstr="update 物品 set 数量=数量+1 where 名称='"&commodityname&"' and 所有者='"&un&"'"
	kl="经过千辛万苦的跋涉，["&un&"]终于抵达了希望的彼岸，成功取得彩虹之神的宝藏，虽然耗费了1体力，却得到一颗神秘宝石，["&un&"]由于过于贫困，找官府换成了白银<img src=data/251.gif>两，精力10000点！"
end if
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
end if
says="<font color=red><b>【抢夺战利品】</b></font>"&kl	
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
		newtalkarr(599)=""&says&""
        newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr
%>
<script Language=Javascript>location.href ='javascript:window.close()';</script>
