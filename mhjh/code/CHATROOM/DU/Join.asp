<!--#include file="data.asp"--> 
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
chatroomsn=session("yx8_mhjh_userchatroomsn")
username=session("yx8_mhjh_username")
shul=trim(Request.Form("zs"))
dx=trim(Request.Form("xz"))
pei=trim(Request.Form("re"))
pai=trim(Request.Form("rg"))
if session("yx8_mhjh_username")<>"" then
if Request.form("ac1")="参加比武" then
   if Application("yx8_mhjh_hong")="无" then

   sql="select * from 参赛 where 场次="&Application("yx8_mhjh_changci")&"+1 and 名称='"&username&"'"
   Set Rs=connt.Execute(sql)
      If rs.EOF Or rs.bof Then
         Application("yx8_mhjh_changci")=Application("yx8_mhjh_changci")+1
         sql="insert into  参赛(名称,场次,结局,是否) values('"&username&"',"&Application("yx8_mhjh_changci")&",'无',False)"
         connt.execute sql
         Application("yx8_mhjh_hong")=username
         session("yx8_mhjh_ok")="ko"
         msg=""&session("yx8_mhjh_username")&",参加第"&Application("yx8_mhjh_changci")&"场比武成功"
         rs.close
      else
         msg="错误，"&session("yx8_mhjh_username")&"你已经参加比武！"
      End If
   else if Application("yx8_mhjh_hei")="无" then

   sql="select * from 参赛 where 场次="&Application("yx8_mhjh_changci")&" and 名称='"&username&"'"
   Set Rs=connt.Execute(sql)
       if rs.EOF or rs.BOF  then
          sql="insert into  参赛(名称,场次,结局,是否) values('"&username&"',"&Application("yx8_mhjh_changci")&",'无',False)"
          connt.execute sql
          msg=""&session("yx8_mhjh_username")&",参加第"&Application("yx8_mhjh_changci")&"场比武成功"
          session("yx8_mhjh_ok")="ko"
          Application("yx8_mhjh_hei")=username
          rs.close
       else
          msg="错误，"&session("yx8_mhjh_username")&"你已经参加比武！"
       end if  
   else
      msg="错误，"&session("yx8_mhjh_username")&",比武席位已经满了，下一场吧"
  end if
  end if
end if


if Request.form("ac2")="我要认输" then



   sql="select * from 参赛 where 场次="&Application("yx8_mhjh_changci")&" and 是否=False and 名称='"&username&"'"
   Set Rs=connt.Execute(sql)
  if rs.EOF or rs.BOF then
   rs.Close
   msg="错误，"&session("yx8_mhjh_username")&"你没有参加该场比武！"
  else
   if username=Application("yx8_mhjh_hong") then 
   session("yx8_mhjh_ok")=""
   sql="update 参赛 set 结局='输',是否=True where 场次="&Application("yx8_mhjh_changci")&" and 名称='"&Application("yx8_mhjh_hong")&"'"
   connt.execute sql
   sql="update 参赛 set 结局='胜',是否=True where 场次="&Application("yx8_mhjh_changci")&" and 名称='"&Application("yx8_mhjh_hei")&"'"
   connt.execute sql
   sql="update 投注 set 胜利者='"&Application("yx8_mhjh_hei")&"' where 场次="&Application("yx8_mhjh_changci")&""
   connt.execute sql
   else if username=Application("yx8_mhjh_hei") then  
   sql="update 参赛 set 结局='输',是否=True where 场次="&Application("yx8_mhjh_changci")&" and 名称='"&Application("yx8_mhjh_hei")&"'"
   connt.execute sql
   sql="update 参赛 set 结局='胜',是否=True where 场次="&Application("yx8_mhjh_changci")&" and 名称='"&Application("yx8_mhjh_hong")&"'"
   connt.execute sql
   sql="update 投注 set 胜利者='"&Application("yx8_mhjh_hong")&"' where 场次="&Application("yx8_mhjh_changci")&""
   connt.execute sql
   session("yx8_mhjh_ok")=""
   end if
   end if
   Application("yx8_mhjh_hong")="无" 
   Application("yx8_mhjh_hei")="无" 
   msg=""&session("yx8_mhjh_username")&"你认输了"
  end if
end if

if Request.form("ac3")="我要下注" then
 if session("yx8_mhjh_ok")="" then
 if dx=Application("yx8_mhjh_hong") then

   sql="select * from 投注 where 场次="&Application("yx8_mhjh_changci")&" and 下注者='"&username&"'"
   Set Rs=connt.Execute(sql)
   if rs.EOF or rs.BOF then
   sql="insert into 投注(下注者,参赛者,数量,赔率,场次,领奖) values('"&username&"','"&dx&"','"&shul&"','"&pei&"',"&Application("yx8_mhjh_changci")&",False)"
   connt.execute sql
   msg=""&session("yx8_mhjh_username")&"，投注成功"
  else
   msg="错误，"&session("yx8_mhjh_username")&"你已经投注"
  end if
 else

   sql="select * from 投注 where 场次="&Application("yx8_mhjh_changci")&" and 下注者='"&username&"'"
   Set Rs=connt.Execute(sql)
   if rs.EOF or rs.BOF then
   sql="insert into 投注(下注者,参赛者,数量,赔率,场次,领奖) values('"&username&"','"&dx&"','"&shul&"','"&pai&"',"&Application("yx8_mhjh_changci")&",False)"
   connt.execute sql
   msg=""&session("yx8_mhjh_username")&"，投注成功"
   else
   msg="错误，"&session("yx8_mhjh_username")&"你已经投注"
   rs.Close
  end if
 end if
 else 
 msg="你是比武者，不能下注！"
 end if
end if
 	    Application.Lock
 	    talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		dim newtalkarr(600) 
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
		newtalkarr(599)="<font color=red>【消息】</font><b><font color=red>"&msg&"</font></b>" 
		newtalkarr(600)=chatroomsn 
		Application("yx8_mhjh_talkarr")=newtalkarr
		Application.UnLock
else
   response.redirect "../../error.asp?id=016"
rs.Close 
set rs=nothing
connt.Close
set connt=nothing
end if
%>
<head><title></title><LINK href="../css3.css" rel=stylesheet></head>
<body oncontextmenu=self.event.returnValue=false topMargin=150 background='../bg1.gif'>
<div align=center><font color=FF0000><%=msg%></font><br><br><a href="anqi.asp">返回</a></div> 
</body> 


