<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
username=session("yx8_mhjh_username")
grade=session("yx8_mhjh_usergrade")
chatroomsn=session("yx8_mhjh_userchatroomsn")
if username="" then Response.Redirect "../error.asp?id=016"
if grade<6 then
Response.Write "<script language=javascript>{alert('赌马需要6级以上\n按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if 
money=abs(int(clng(Request.form("money"))))
if money<10000 or money>1000000 then
Response.Write "<script language=javascript>{alert('所输入得数字不符合1-100万\n按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if 
radiobutton=Request.form("radiobutton")
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"chatroom")=0 then 
Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if 
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
        response.write "禁止站外提交！"
        response.end
end if
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="select 银两 FROM 用户 WHERE 姓名='" & username &"'"
set rs=conn.execute(sql)
if rs("银两")<money then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你钱好象不够！');parent.history.go(-1);}</script>"
	Response.End
end if
	sql="Select count(*) As 数量 from d where c='马'"
	set rs=conn.execute(sql)
	ds=rs("数量")
	if ds>10 then
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('["&durs&"]现在开局，稍后再下注！');parent.history.go(-1);}</script>"
		Response.End 		
	end if
sql="select top 1 * FROM d WHERE c='马' and xm='"& username &"'"
set rs=conn.execute(sql)
if not(rs.eof) or not(rs.bof) then
	temp="提示："&username&"你买了["&rs("ydx")&"]号马赢,计："&rs("yl")&"两，等着开始吧!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('"&temp&"');parent.history.go(-1);}</script>"
	Response.End 
else
	conn.execute "update 用户 set 银两=银两-"&money&" where 姓名='" & username &"'"
	conn.execute "insert into d(xm,yl,ydx,c) values ('"&username&"',"&money&",'"&radiobutton&"','马')"	
    tmprs=conn.execute("Select count(*) As 数量 from d where c='马'")
    dars=tmprs("数量")
if dars>=10 then
	randomize timer
	mmm=int(rnd*4)+1
	yinma=""
	sql="select * FROM d WHERE c='马' and ydx='"&mmm&"'"
	set rs=conn.execute(sql)
	do while not rs.bof and not rs.eof
		yinma=yinma&"["&rs("xm")&" "& rs("yl") &"] "
		conn.execute "update 用户 set 银两=银两+"&rs("yl")*2&",体力=体力+1000000,内力=内力+1000000 where 姓名='"& rs("xm") &"'"
		rs.movenext
	loop
	conn.execute "delete * from d where c='马'"
	says="<font color=blue><b>[马赛开始]</b></font>紧张激烈的马赛开始了。只见5匹高头竣马快速奔跑………最终<font color=red>"&mmm&"</font>号码<img src=horse/"&mmm&".gif>跑在了最前面，得到了第一名，玩家：<font color=blue><b>"&yinma&"</b></font>玩家幸运押中,每人得到下注的<font color=red>2</font>倍的银两,所有押中玩家内力、体力各加<font color=red>+100万</font>！"
	call chat(says)
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('[操作完成]');parent.history.go(-1);}</script>"	
	Response.End 
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=blue><b>[江湖马场]</b></font>"&username&"在江湖马上选了一匹好马<img src=horse/"&radiobutton&".gif>["&radiobutton&"]号，在这匹马上下注:<font color=red>"&money&"</font>两，现在还差<font color=blue>"&(10-dars)&"</font>开赛！赢了你可就发财了，哈哈"
call chat(says)
Response.Write "<script language=JavaScript>{alert('[操作完成]');parent.history.go(-1);}</script>"	
Response.End 
end if
sub chat(says)
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
        newtalkarr(600)=chatroomsn
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr	
end sub
%>