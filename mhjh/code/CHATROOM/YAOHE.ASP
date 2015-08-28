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
if grade<10 then
Response.Write "<script language=javascript>{alert('坐庄需要10级以上\n按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if
    set conn=server.createobject("adodb.connection") 
    conn.open Application("yx8_mhjh_connstr")
    set rs=server.CreateObject ("adodb.recordset")
	sql= "select xm,sf,sj FROM d WHERE xm='"&un&"' and sf='庄家'"
	set rs=conn.execute(sql)
	if rs.EOF or rs.BOF then
	   rs.close
	   set rs=nothing	
	   conn.close
	   set conn=nothing
	   Response.Write "<script language=JavaScript>{alert('您不是庄家，无法吆喝！');parent.history.go(-1);}</script>"
	   Response.End
	else
	    sj=DateDiff("s",rs("sj"),now())
        if sj<60 then
        ss=60-sj
       Response.Write "<script language=JavaScript>{alert('想刷屏啊，请你等上["& ss &"]秒再吆喝！');parent.history.go(-1);}</script>"
	   Response.End
	else
	sql= "Select count(*) As 数量 from d where yl<>0 and c='银'"
	set rs=conn.execute(sql)
	durs=rs("数量")
	conn.execute "update d set sj=now() where xm='"&un&"' and sf='庄家' and c='银'"
    says="<b><font color=FF0000>【在线赌博】庄家:<font color=#000000>["&un&"]</font>叫到：买得多赢得多，买得少赢得少！买了您后悔，不买您更后悔，还有:<font color=#000000>"& (10-durs) &"</font>个下注位开局……</font></b><img src=../chatroom/image/sz.gif>"
    call chat(says)	
	end if
	end if
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('[操作成功]');parent.history.go(-1);}</script>"	
Response.End 
    
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


