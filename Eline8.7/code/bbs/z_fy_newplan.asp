<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_fy_Conn.asp"-->
<!--#include file="z_fy_const.asp"-->
<%
'=========================================================
' Plusname: 法院监狱第三版
' File: z_fy_newplan.asp
' For Version: 6.0sp2
' Date: 2003-2-10
' Script Written by Hero20000
'=========================================================
if request.form("topic")="" or request.form("text")="" then
%>
  <script language=vbscript>
    MsgBox "请填写投诉的理由和举报内容！"
    location.href = "javascript:history.back()"
  </script>
<%
else
  dim checkad,sqlad
  dim topic,play,text
  towho=replace(request.form("towho"),"'","")
  sqlad="select userclass from [user] where username='"&towho&"' "
  set checkad=conn.execute(sqlad)
  if  checkad.eof then
  %>
    <script language=vbscript>
     MsgBox "对不起，您投诉的人不存在！"
     location.href = "javascript:history.back()"
   </script>
  <%
     response.end
  else
   if  checkad(0)="管理员" then
   %>
       <script language=vbscript>
      MsgBox "对不起，您不能投诉管理员！"
      location.href = "javascript:history.back()"
       </script>
    <%
       response.end
   end if
 end if
 topic=replace(request.form("topic"),"'","")
 play=replace(request.form("play"),"'","")
 text=replace(request.form("text"),"'","")
 dim fysql,yugao
 yugao=membername
 connfy.execute("insert into fy(title,bg,ygtext,yg,ygyq,dateandtime) values ('"&topic&"','"&towho&"','"&text&"','"&yugao&"','"&play&"',now())")
 content="社区法院已受理"&membername&"的诉状。投诉您:"&topic&"-"&text&"。要求对您进行 "&play&" 的处罚。请于2日内到社区法院申诉，超期将缺席审判。"
 call messto(towho,"社区法院","法院传票",content)
 response.redirect("z_fy_over.asp")
end if
%>