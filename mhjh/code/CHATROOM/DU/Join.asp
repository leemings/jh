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
if Request.form("ac1")="�μӱ���" then
   if Application("yx8_mhjh_hong")="��" then

   sql="select * from ���� where ����="&Application("yx8_mhjh_changci")&"+1 and ����='"&username&"'"
   Set Rs=connt.Execute(sql)
      If rs.EOF Or rs.bof Then
         Application("yx8_mhjh_changci")=Application("yx8_mhjh_changci")+1
         sql="insert into  ����(����,����,���,�Ƿ�) values('"&username&"',"&Application("yx8_mhjh_changci")&",'��',False)"
         connt.execute sql
         Application("yx8_mhjh_hong")=username
         session("yx8_mhjh_ok")="ko"
         msg=""&session("yx8_mhjh_username")&",�μӵ�"&Application("yx8_mhjh_changci")&"������ɹ�"
         rs.close
      else
         msg="����"&session("yx8_mhjh_username")&"���Ѿ��μӱ��䣡"
      End If
   else if Application("yx8_mhjh_hei")="��" then

   sql="select * from ���� where ����="&Application("yx8_mhjh_changci")&" and ����='"&username&"'"
   Set Rs=connt.Execute(sql)
       if rs.EOF or rs.BOF  then
          sql="insert into  ����(����,����,���,�Ƿ�) values('"&username&"',"&Application("yx8_mhjh_changci")&",'��',False)"
          connt.execute sql
          msg=""&session("yx8_mhjh_username")&",�μӵ�"&Application("yx8_mhjh_changci")&"������ɹ�"
          session("yx8_mhjh_ok")="ko"
          Application("yx8_mhjh_hei")=username
          rs.close
       else
          msg="����"&session("yx8_mhjh_username")&"���Ѿ��μӱ��䣡"
       end if  
   else
      msg="����"&session("yx8_mhjh_username")&",����ϯλ�Ѿ����ˣ���һ����"
  end if
  end if
end if


if Request.form("ac2")="��Ҫ����" then



   sql="select * from ���� where ����="&Application("yx8_mhjh_changci")&" and �Ƿ�=False and ����='"&username&"'"
   Set Rs=connt.Execute(sql)
  if rs.EOF or rs.BOF then
   rs.Close
   msg="����"&session("yx8_mhjh_username")&"��û�вμӸó����䣡"
  else
   if username=Application("yx8_mhjh_hong") then 
   session("yx8_mhjh_ok")=""
   sql="update ���� set ���='��',�Ƿ�=True where ����="&Application("yx8_mhjh_changci")&" and ����='"&Application("yx8_mhjh_hong")&"'"
   connt.execute sql
   sql="update ���� set ���='ʤ',�Ƿ�=True where ����="&Application("yx8_mhjh_changci")&" and ����='"&Application("yx8_mhjh_hei")&"'"
   connt.execute sql
   sql="update Ͷע set ʤ����='"&Application("yx8_mhjh_hei")&"' where ����="&Application("yx8_mhjh_changci")&""
   connt.execute sql
   else if username=Application("yx8_mhjh_hei") then  
   sql="update ���� set ���='��',�Ƿ�=True where ����="&Application("yx8_mhjh_changci")&" and ����='"&Application("yx8_mhjh_hei")&"'"
   connt.execute sql
   sql="update ���� set ���='ʤ',�Ƿ�=True where ����="&Application("yx8_mhjh_changci")&" and ����='"&Application("yx8_mhjh_hong")&"'"
   connt.execute sql
   sql="update Ͷע set ʤ����='"&Application("yx8_mhjh_hong")&"' where ����="&Application("yx8_mhjh_changci")&""
   connt.execute sql
   session("yx8_mhjh_ok")=""
   end if
   end if
   Application("yx8_mhjh_hong")="��" 
   Application("yx8_mhjh_hei")="��" 
   msg=""&session("yx8_mhjh_username")&"��������"
  end if
end if

if Request.form("ac3")="��Ҫ��ע" then
 if session("yx8_mhjh_ok")="" then
 if dx=Application("yx8_mhjh_hong") then

   sql="select * from Ͷע where ����="&Application("yx8_mhjh_changci")&" and ��ע��='"&username&"'"
   Set Rs=connt.Execute(sql)
   if rs.EOF or rs.BOF then
   sql="insert into Ͷע(��ע��,������,����,����,����,�콱) values('"&username&"','"&dx&"','"&shul&"','"&pei&"',"&Application("yx8_mhjh_changci")&",False)"
   connt.execute sql
   msg=""&session("yx8_mhjh_username")&"��Ͷע�ɹ�"
  else
   msg="����"&session("yx8_mhjh_username")&"���Ѿ�Ͷע"
  end if
 else

   sql="select * from Ͷע where ����="&Application("yx8_mhjh_changci")&" and ��ע��='"&username&"'"
   Set Rs=connt.Execute(sql)
   if rs.EOF or rs.BOF then
   sql="insert into Ͷע(��ע��,������,����,����,����,�콱) values('"&username&"','"&dx&"','"&shul&"','"&pai&"',"&Application("yx8_mhjh_changci")&",False)"
   connt.execute sql
   msg=""&session("yx8_mhjh_username")&"��Ͷע�ɹ�"
   else
   msg="����"&session("yx8_mhjh_username")&"���Ѿ�Ͷע"
   rs.Close
  end if
 end if
 else 
 msg="���Ǳ����ߣ�������ע��"
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
		newtalkarr(595)="���" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)="<font color=red>����Ϣ��</font><b><font color=red>"&msg&"</font></b>" 
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
<div align=center><font color=FF0000><%=msg%></font><br><br><a href="anqi.asp">����</a></div> 
</body> 


