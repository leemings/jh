<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_fy_Conn.asp"-->
<!--#include file="z_fy_const.asp"-->
<%
'=========================================================
' Plusname: ��Ժ����������
' File: z_fy_newplan.asp
' For Version: 6.0sp2
' Date: 2003-2-10
' Script Written by Hero20000
'=========================================================
if request.form("topic")="" or request.form("text")="" then
%>
  <script language=vbscript>
    MsgBox "����дͶ�ߵ����ɺ;ٱ����ݣ�"
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
     MsgBox "�Բ�����Ͷ�ߵ��˲����ڣ�"
     location.href = "javascript:history.back()"
   </script>
  <%
     response.end
  else
   if  checkad(0)="����Ա" then
   %>
       <script language=vbscript>
      MsgBox "�Բ���������Ͷ�߹���Ա��"
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
 content="������Ժ������"&membername&"����״��Ͷ����:"&topic&"-"&text&"��Ҫ��������� "&play&" �Ĵ���������2���ڵ�������Ժ���ߣ����ڽ�ȱϯ���С�"
 call messto(towho,"������Ժ","��Ժ��Ʊ",content)
 response.redirect("z_fy_over.asp")
end if
%>