<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
un=session("yx8_mhjh_username")
chatroomsn=session("yx8_mhjh_userchatroomsn")
usercorp=session("yx8_mhjh_usercorp")
if un="" then Response.Redirect "../../error.asp?id=016"
if usercorp<>"出家" then
Response.write "<script language='javascript'>alert ('你不是出家人,还想超渡？');window.close();</script>"
Response.end
end if
xm=request("id")
mc=request.querystring("shihe")
if un=xm then
Response.write "<script language='javascript'>alert ('自己给自己超度啊？？');window.close();</script>"
Response.end
end if
if not(mc="我佛慈悲" or mc="佛光普照" or mc="万佛朝宗") then Response.Redirect "../../error.asp?id=024"
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="select 姓名,等级,银两,积分 from 用户 where 形态='鬼魂' and 姓名='"&xm&"'"
rst.open sql,conn,1,1
if rst.eof or rst.bof then
Response.write "<script language='javascript'>alert ('你是不是弄错了，"&xm&"不是鬼魂啊');location.href='cd.asp';</script>"
Response.end
else
dj=rst("等级") 
yl=rst("银两")
jf=rst("积分")
if yl<500000 then 
Response.write "<script language='javascript'>alert ('这鬼魂生前是个穷鬼，连50万也没有，做法事的钱都没了');location.href='cd.asp';</script>"
Response.end
end if
rst.Close  
end if
%><!--#include file="data.asp"--><%
sql="select 等级,攻击,防御,类型,时间 from mj where 所有者='"&un&"' and 名称='"&mc&"'"
set rs=connb.execute(sql)
if rs.EOF or rs.BOF then
Response.write "<script language='javascript'>alert ('你是不是弄错了，你没有练过"&mc&"啊');location.href='cd.asp';</script>"
Response.end
else
dj2=rs("等级")
gj=rs("攻击")
fy=rs("防御")
rs.Close 
randomize timer
sz=int(rnd*9)+1
if dj2<dj and sz>8 then
conn.Execute "update 用户 set 体力=体力-"&dj*100&",银两=银两-500000,形态='人' where 姓名='"&xm&"'"
conn.Execute "update 用户 set 资质=资质+"&sz&",银两=银两+100000 where 姓名='"&un&"'"
sql="update mj set 攻击=攻击-"&sz*100&",防御=防御-"&dj*10&" where 所有者='"&un&"' and 名称='"&mc&"'"
Set Rs=connb.Execute(sql)
msg="超渡成功,禅功法力减少"&sz*100&",魔阿力减少"&dj2*10&",资质增加"&sz&",银两增加100000"
elseif  dj2>dj then
conn.Execute "update 用户 set 体力=体力-"&dj*100&",银两=银两-500000,形态='人' where 姓名='"&xm&"'"
conn.Execute "update 用户 set 资质=资质+"&sz&",银两=银两+100000 where 姓名='"&un&"'"
sql="update mj set 攻击=攻击+"&sz*100&",防御=防御+"&(10-sz)*dj&" where 所有者='"&un&"' and 名称='"&mc&"'"
Set Rs=connb.Execute(sql)
msg="超渡成功,禅功法力增加"&sz*100&",魔阿力增加"&(10-sz)*dj2&",资质增加"&sz&",银两增加100000"
else
if jf>(gj+fy)*10 then
conn.Execute "update 用户 set 体力=体力-"&dj*100&",银两=银两-500000,状态='死亡' where 姓名='"&un&"'"
sql="update mj set 攻击=攻击-"&sz*dj2&",防御=防御-"&(10-sz)*dj2&" where 所有者='"&un&"' and 名称='"&mc&"'"
Set Rs=connb.Execute(sql)
session.Abandon
Response.write "<script language='javascript'>alert ('可怜你的禅功不敌鬼魂的催命功,被杀死了');location.href='../../exit.asp';</script>"
Response.end
else
conn.Execute "update 用户 set 资质=资质-"&sz&",银两=银两-500000 where 姓名='"&un&"'"
sql="update mj set 攻击=攻击-"&sz*100&",防御=防御-"&(10-sz)*dj2&" where 所有者='"&un&"' and 名称='"&mc&"'"
msg="超渡失败,禅功法力增加"&sz*100&",魔阿力增加"&(10-sz)*dj2&",资质减少"&sz&",银两减少500000"
end if
end if 
end if 
set rst=nothing 
conn.Close              
set conn=nothing
set rs=nothing 
connb.Close              
set connb=nothing
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
		newtalkarr(599)="【超渡亡魂】"&un&""&msg&""
        newtalkarr(600)=chatroomsn
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr	    
%> 
<html>
<head>
<title>超渡</title>
<link rel=stylesheet href='../../style.css'>
</head>
<body oncontextmenu=self.event.returnValue=false background='../../chatroom/bg1.gif'>
　
<div align=center>
<font color=0000ff size=4>超渡亡魂</font>
<hr>
</div>
<p align="center"><%=msg%><br><br><input type=button value='返回' onclick="javascript:location.href='cd.asp'"></p>
</body>
</html>

