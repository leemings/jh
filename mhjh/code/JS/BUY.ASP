<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
commodityid=Request.form("id")
sl=int(abs(Request.form("sl")))
if sl<1 or sl>9 then Response.Response.Redirect "../error.asp?id=520"
if InStr(id,"oR")<>0 or InStr(id,"Or")<>0 or InStr(id,"OR")<>0 or InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"　")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../error.asp?id=54"
if isnumeric(commodityid) then
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select * from 商店 where id="&commodityid
rst.Open sqlstr,conn
if rst.EOF or rst.BOF then
msg="错误，不出售此物品！"
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
rst.Open "select 银两 from 用户 where 姓名='"&username&"' and 银两>="&commodityprice*sl,conn
if rst.EOF or rst.BOF then
msg="错误，你的钱好象没有带够哟！"
else
rst.Close
sqlstr="select * from 物品 where 名称='"&commodityname&"' and 所有者='"&username&"'"
rst.Open sqlstr,conn
if rst.EOF or rst.BOF then
sqlstr="insert into 物品(名称,属性,体力,内力,攻击,防御,特效,价格,数量,所有者,寄售,装备) values('"&commodityname&"','"&commoditytype&"',"&commodityhealth&","&magic&","&attack&","&defence&",'"&especial&"',"&commodityprice/2&","&sl&",'"&username&"',False,False)"
	kl="<font color=0000FF>"&username&"</font>购买了"&sl&"个"&commodityname&"！"
else
sqlstr="update 物品 set 数量=数量+" & sl & " where 名称='"&commodityname&"' and 所有者='"&username&"'"
	kl="<font color=0000FF>"&username&"</font>购买了"&sl&"个"&commodityname&"！"
end if
conn.BeginTrans
conn.Execute sqlstr
conn.Execute "update 用户 set 银两=银两-"&commodityprice*sl&" where 姓名='"&username&"'"
conn.CommitTrans
msg="承惠纹银"&commodityprice*sl&"两，这是"&sl&"个"&commodityname&"您收好了！"
end if
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
says="<font color=red><b>【商店采购】</b></font>"&kl	
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
else
msg="错误，不出售此物品！"
end if
%>
<head><title>商店</title><LINK href="../style.css" rel=stylesheet><script language=javascript>setTimeout('history.back();',3000);</script></head>
<body background='../chatroom/bg1.gif' ncontextmenu=self.event.returnValue=false topMargin=200>
<div align=center>3秒钟后自动返回<br>
<font color=FF0000>&nbsp; <%=msg%></font></div>
<p align=center><a href="#" onclick=javascript:history.back(); onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回</a></p>
</body>
