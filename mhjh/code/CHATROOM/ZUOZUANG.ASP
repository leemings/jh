<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
un=session("yx8_mhjh_username")
grade=session("yx8_mhjh_usergrade")
chatroomsn=session("yx8_mhjh_userchatroomsn")
if un="" then  Response.Redirect "../error.asp?id=016"
if grade="" then  Response.Redirect "../error.asp?id=016"
if grade<6 then
Response.Write "<script language=javascript>{alert('赌博需要6级以上\n按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if
    set conn=server.createobject("adodb.connection") 
    conn.open Application("yx8_mhjh_connstr")
    set rs=server.CreateObject ("adodb.recordset")	
	sql= "select 姓名,银两 FROM 用户 WHERE 姓名='"&un&"' and 银两>100000"
	set rs=conn.execute(sql)
	if rs.EOF or rs.BOF then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('你的银两不够10万，不能作庄！');parent.history.go(-1);}</script>"
		Response.End 
	else
	sql= "select top 1 xm FROM d WHERE sf='庄家'"
	set rs=conn.execute(sql)
	if not(rs.eof) or not(rs.bof) then
		Response.Write "<script language=JavaScript>{alert('现在["&rs("xm")&"]为庄家,你不能作庄！！');parent.history.go(-1);}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End 
	end if
	conn.execute "insert into d(xm,sf,yl,ydx,sj,c) values ('"&un&"','庄家',0,'庄',now(),'银')"
	conn.execute "update 用户 set 银两=银两-100000 where 姓名='"&un&"'"
	zuo="<b><font color=FF0000>【赌博作庄】<font color=#000000>["&un&"]</font>在赌场申请作庄成功，现在大家可以豪赌一把了,快快来下注呀……</font></b><img src=../chatroom/image/sz.gif>"
	call chat(says)
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('[操作成功]');parent.history.go(-1);}</script>"	
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
		newtalkarr(599)=""&zuo&""
        newtalkarr(600)=chatroomsn
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr
   end sub
%>

