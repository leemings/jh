<!--#include file="conn.asp"-->
<!--#include file="z_fy_Conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_fy_const.asp"-->
<%
'=========================================================
' Plusname: ��Ժ����������
' File: z_fy_looktj.asp
' For Version: 6.0sp2
' Date: 2003-2-10
' Script Written by Hero20000
'=========================================================
dim fanren,fysql,fyrs
fanren=checkstr(request("name"))
call getfyconfig()
%>
<html>
<head>
<title>�鿴<%=fanren%>���ԣ��¼���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body background="Images/img_fy/Bg001.gif" >
<!-- #include file="inc/FORUM_CSS.asp" -->    
<table border=0  cellspacing=1 cellpadding=3 align="center" width=60% class=tableborder1>
<%
select case  checkstr(request("action"))
case "list" 
  call listh()
case "delmy"
  call delmy()
case "delall"
  call delall()
case else
if log_set(0)="1" then
fysql="delete from log where class='����' and tousername='"&fanren&"' and title<>'����' and id not in (select top "&clng(log_set(1))&" id from log where class='����' and tousername='"&fanren&"' and title<>'����' order by id desc)"
connfy.execute(fysql)
end if
  fysql="SELECT username,title,content,dateandtime FROM log WHERE class='����' and tousername='"&fanren&"' order by dateandtime desc"
  set fyrs=connfy.execute(fysql)
  if fyrs.eof then
    response.write "<tr><td class=tablebody1 align=center>�Ǻǣ��������¼�/������û����Ҫ�� "&fanren&" ��¼Ŷ��</td></tr>"
  else
    do while Not fyRS.Eof 
     response.write "<tr><td  height=24 width=40% class=tablebody2>���ڣ�"&fyrs(3)&"</td>"
     response.write "<td height=24 class=tablebody2 >�¼���"&fyrs(1)&"</td>"
     response.write "<td height=24 width=40% class=tablebody2 >"
         select case fyrs(1)
           case "����"
        response.write "����Ա/���٣�"
      case "����"
        response.write "����Ա/���٣�"
      case "����"
        response.write "���ˣ�"
      case "����"
        response.write "ǿ����"
      case "̽��"
        response.write "�����ˣ�"
    end select 
    if fyrs(1)="����" and qj_set(3)<>"1" then
      response.write " ������</td></tr>"
    else  
    response.write " "&fyrs(0)&"</td></tr>"
    end if
     response.write "<tr><td class=tablebody1 colspan=3 >���飺<br>"&fyrs(2)&"<br></td></tr>"
    fyRS.MoveNext
    Loop
  end if
end select

%>
</table>
 
<br>
<div align="center">
<%
if membername=fanren or master or jymaster then 
response.write "��<a href=z_fy_looktj.asp?action=delmy&name="&fanren&">"
 if master or jymaster then 
  response.write "��ո��û�������¼</a>��<br>"
 else
   response.write "����ҵļ���</a>��<br>"
 end if
end if
if master then response.write  "��<a href=z_fy_looktj.asp?action=delall>����Ա������м���</a>��<br>"
%><input type="button" value="�ر�" onClick="window.close()" name="button">
</div>
<% 
sub delmy()   
 if membername<>checkstr(request("name")) and  (not master) and (not jymaster) then   
  response.write "<tr><td class=tablebody1>�Բ�����ɾ���˵ļ�¼��ʲô��</td></tr>"   
 else
  if membername=checkstr(request("name")) then   
     fysql="delete * from log where class='����' and title<>'����' and tousername='"&membername&"'"
  else
     fysql="delete * from log where class='����' and tousername='"&checkstr(request("name"))&"'"
  end if
  connfy.execute(fysql)   
  Response.Redirect "z_fy_looktj.asp"   
 end if   
end sub   
   
sub delall()   
 if not master then   
   response.write "<tr><td class=tablebody1>�Բ����㲻�ǹ���Աѽ��</td></tr>"   
 else   
   fysql="delete * from log where class='����' "   
   connfy.execute(fysql)   
   Response.Redirect "z_fy_looktj.asp"   
 end if   
end sub   
   
sub listh()   
  fysql="SELECT DISTINCT tousername FROM log WHERE class='����' and (title='����' or title='����')"   
  set fyrs=connfy.execute(fysql)   
  if fyrs.eof then   
    response.write "<tr><td class=tablebody1 align=center>�Ǻǣ�û�м�¼Ŷ��</td></tr>"   
  else   
    dim unum   
    unum=0   
    response.write "<tr>"   
    do while Not fyRS.Eof  
    unum=unum+1   
     response.write "<td class=tablebody1 align=center>��<a href=z_fy_looktj.asp?name="&fyrs(0)&" title='����鿴����'>"&fyrs(0)&"</a>��</td>"   
    if (unum mod 4=0) then response.write "</tr><tr>"   
    fyRS.MoveNext   
    Loop   
    response.write "</tr>"   
  end if   
end sub   
  
 %>   
</body></html>