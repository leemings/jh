<%
id=request("id")
my=session("yx8_mhjh_username")
shihe=request.querystring("shihe")
%><!--#include file="dadata.asp"-->
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="select * from 用户 where 姓名='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
	response.write "你不是江湖中人，不能订购酒宴"
	conn.close
	response.end
else

sql="SELECT * FROM 宴会 where ID=" & id 
	Set Rs=connt.Execute(sql)
        wu=rs("宴会名")
	yin=rs("售价")
	lx=rs("类型")
        nl=rs("内力")
        tl=rs("体力")
        jb=rs("资质")
        sl=rs("数量")%>

<%
		sql="select * from 用户 where 姓名='" & my & "'"
		rs=conn.execute(sql)
		if yin<=rs("银两") then
			sql="select * from 宴会列表 where 宴会名='" & wu & "' and 拥有者='" & my & "'"
			set rs=connt.execute(sql)
			if rs.eof or rs.bof then
                                   sql="update 用户 set 银两=银两-" & yin & " where 姓名='" & my & "'"
			rs=conn.execute(sql)
			sql="insert into 宴会列表(宴会名,拥有者,类型,内力,体力,资质,数量,售价,时间,说明) values ('"&wu&"','"&my&"','"&lx&"',"&nl&","&tl&","&jb&","&sl&","&yin&",now(),'"&shihe&"')"
			rs=connt.execute(sql)
            
			connt.close
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
		newtalkarr(599)="<font color=red>【江湖消息】</font><font color=red>"&my&"在魔幻大酒店举行盛大的※"&wu&"※宴会,参加者可以增加内力"&nl&",增加体力"&tl&",增加资质"&jb&" ,各位大侠别辜负了人家的一片好意,赶紧去赴宴啊</font>" 
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
			Response.Redirect "jd.asp"
			else
            
				
					response.write "不能定酒宴，原因：你已定了同样规格的酒席！"
					connt.close
					response.end
				
			end if
		else
			response.write "不能定酒宴，原因：你的银两不够了"
			connt.close
			response.end
		end if
   rs.close
   set rs=nothing
end if
%>