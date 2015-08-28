<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=Session("yx8_mhjh_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rs=server.CreateObject("adodb.recordset")
if Application("yx8_mhjh_klt19")=0 then
	Response.Write "<script Language=Javascript>alert('提示：怪物已经不在了！');</script>"
	Response.Write "<script Language=Javascript>location.href = '../option.asp';</script>"
	Response.end
end if
tempjs=int(abs(clng(Application("yx8_mhjh_klt19"))))
'if tempjs<>r then
	'Response.Write "<script Language=Javascript>alert('提示：滚，你想搞什么，按回车键不行了!');</script>"
	'Response.Write "<script Language=Javascript>location.href = '../option.asp';</script>"
	'response.end
'end if
conn.execute "update 用户 set 体力=体力-"&Application("yx8_mhjh_klt19")&" where 姓名='"&username&"'"
rs.open "select 体力,门派 FROM 用户 WHERE 姓名='"&username&"'",conn
if rs("门派")="十八地狱" then
	Response.Write "<script Language=Javascript>alert('提示：你小子是不是没事找事呀，鬼魂还想得便宜啊！');</script>"
	Response.Write "<script Language=Javascript>location.href = '../option.asp';</script>"
	Application("yx8_mhjh_klt19")=0
	Response.end
end if
if rs("体力")<0 then
	conn.execute "update 用户 set 状态='死亡',被杀=被杀+1 where 姓名='"&username&"'"
conn.execute "insert into 英烈堂(时间,死者,凶手,死因) values(now(),'"&username&"','吝啬','体力不足累死')"
	kl="<img src='data/kl.gif'>太不幸了，["&username&"]为了抵御怪物袭击，舍身取义，把生命永远的留在了彩虹之下，陪伴他的只有千百年仍在变幻着神奇光芒的传说………"
	Response.Write "<script Language=Javascript>alert('提示：你英勇可嘉，但是体力太弱，丢了小命，去复活吧！没能耐别乱闯！');</script>"
	session.Abandon
	response.write "<script language=javascript>parent.top.location.href='../chaterror.asp?id=009';</script>"
else
randomize
r=120 
id=int((r-1)*rnd)+1 
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rs = server.createobject("adodb.recordset") 
rs.open "select * from 商店 where id="&id,conn
if rs.EOF or rs.BOF then
Response.Write "<script Language=Javascript>alert('很不幸，你什么都没得到！');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
Application("yx8_mhjh_klt19")=0
	Response.end
end if
ming=rs("名称")
rs.Close
rs.Open "select * from 物品 where 所有者='"&username&"' and 名称='"&ming&"'",conn
If Rs.Bof OR Rs.Eof then
rs.close 
rs.open "select * from 商店 where 名称='"&ming&"'",conn
conn.Execute "insert into 物品(名称,属性,体力,内力,攻击,防御,特效,数量,价格,所有者,寄售,装备) values('"&ming&"','"&rs("属性")&"',"&rs("体力")&","&rs("内力")&","&rs("攻击")&","&rs("防御")&",'"&rs("特效")&"',1,"&rs("价格")\2&",'"&username&"',false,false)"
	kl="<font color=#0000FF>"&username&"</font>成功打退了怪物,并获得怪物身上的"&ming&"！"
Application("yx8_mhjh_klt19")=0
else
         conn.Execute "update 物品 set 数量=数量+1 where 所有者='"&username&"' and 名称='"&ming&"'"
	kl="<font color=#0000FF>"&username&"</font>成功打退了怪物,并获得怪物身上的"&ming&"！"
	Application.Lock
	Application("yx8_mhjh_klt19")=0
	Application.UnLock
	rs.close
set rs=nothing
conn.close
set conn=nothing
end if
end if
says="<font color=red><b>【打怪成功】</b></font>"&kl	
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
<script Language=Javascript>location.href = "../option.asp";</script>
