<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
if session("yx8_mhjh_inchat")="" then
Response.write "<script language='javascript'>alert ('你不能进入,请先进入聊天室再来。谢谢合作'); window.close();</script>"
Response.End 
        end if
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="select 银两,积分,等级,资质,任务名称 from 用户 where 姓名='"&username&"'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then Response.Redirect "../error.asp?id=016"
jf=rs("积分")
zz=rs("资质")
yl=rs("银两")
dj=rs("等级")
if dj<100 then
mess=""&username&"你还不够100级啊,怎么搞的,跑这里来干什么,回去练练再说吧！"
elseif jf<4000000 then
mess=""&username&"你积分不到400万,目前还无法实现转生！"
elseif yl<100000000 then
mess=""&username&"你的钱不够，都这么高级别了,连这点钱都攒不下，别让人说你白活这么大哦！"
elseif zz<10000 then
mess=""&username&"你资质太差,至少要达到100万的水平,这里不能接受你这样弱智的人！"
else
yl1=yl-100000000
dd=rs("资质")-10000
if rs("任务名称")="" or rs("任务名称")="无" then
mess="恭喜，"&username&"已经实现了转生，级别和积分都变成0，但好处多多哦，以后自己的路自己走吧<br><br>请退出后重新登陆"
conn.execute "update 用户 set 银两="&yl1&",资质="&dd&",积分=0,等级=0,任务名称='转生人',体力加=5000000,内力加=500000 where 姓名='"&username&"'"
elseif rs("任务名称")="转生人" then
mess="恭喜，"&username&"已经实现了由转生人升华为强化人，级别和积分都变成0，但好处多多哦，以后自己的路自己走吧<br><br>请退出后重新登陆"
conn.execute "update 用户 set 银两="&yl1&",资质="&dd&",积分=0,等级=0,任务名称='强化人',体力加=10000000,内力加=1000000 where 姓名='"&username&"'"
elseif rs("任务名称")="强化人" then
mess="恭喜，"&username&"已经实现了由强化人升华为觉醒人，级别和积分都变成0，但好处多多哦，以后自己的路自己走吧<br><br>请退出后重新登陆"
conn.execute "update 用户 set 银两="&yl1&",资质="&dd&",积分=0,等级=0,任务名称='觉醒人',体力加=15000000,内力加=1500000 where 姓名='"&username&"'"
elseif rs("任务名称")="觉醒人" then
mess="恭喜，"&username&"已经实现了由觉醒人升华为真*无敌，级别和积分都变成0，但好处多多哦，以后自己的路自己走吧<br><br>请退出后重新登陆"
conn.execute "update 用户 set 银两="&yl1&",资质="&dd&",积分=0,等级=0,任务名称='真*无敌',体力加=20000000,内力加=2000000 where 姓名='"&username&"'"
end if
end if
rs.Close             
set rs=nothing             
conn.Close             
set conn=nothing       
says="<font color=red><b>【绝世转生】</b></font>"&mess	
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
<script Language=Javascript>alert('<%=mess%>！');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>
