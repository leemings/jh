<!--#include file="conn.asp"-->
<!--#include file="z_fy_Conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_fy_const.asp"-->
<!-- #include file="inc/FORUM_CSS.asp" --> 

<%
'=========================================================
' Plusname: ��Ժ����������
' File: z_fy_tanjian.asp
' For Version: 6.0sp2
' Date: 2003-2-10
' Script Written by Hero20000
'=========================================================
if not founduser then 
  Errmsg=Errmsg+"<br>"+"<li>�οͲ���̽��������"
   call dvbbs_error()
else
  call getfyconfig()
  dim fanren,jailsj,sjpic,sjstr
  call checkjail()
  if not founderr then
    select case checkstr(request("action"))
      case "save" 
        call savetj()
      case "qiang"
        call qiang()       
      case else
        call main()
    end select
  end if
end if
rs.close
set rs=nothing

sub main()  
%>
<html>
<head>
<title><%=membername%>̽������<%=fanren%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body background="Images/img_fy/bg001.gif" >
<form action=z_fy_tanjian.asp?action=save&name=<%=fanren%> method=post>
<table border=0  cellspacing=1 cellpadding=3 align="center" width="60%" ><center>
  <tr> 
    <td align=center > 
<%
randomize 
jailsj=Int((15*rnd)+1)
select case jailsj
 case 1
  sjpic="Images/img_fy/sj1.gif"
  sjstr="����������ʹ�����飬����Լ���������Ϊ��"
 case 2
  sjpic="Images/img_fy/sj2.gif"
  sjstr="����ʼ�μ���ѵ�࣬ѧϰ�������ɷ��档"
 case 3
 sjpic="Images/img_fy/sj3.gif"
  sjstr="���Ѿ���ʼ��ѧ�����µ��书��Ҫ��׼���������Ż�������"
 case 4
 sjpic="Images/img_fy/sj4.gif"
  sjstr="���������������죬����������һ�졣"
 case 5
 sjpic="Images/img_fy/sj5.gif"
  sjstr="���������Ѽ���������ײ���ˣ�������������̹�˲Ž����Ʒ���"
 case 6
 sjpic="Images/img_fy/sj6.gif"
  sjstr="�ﷸ��ͼԽ������������ץ�أ������ڸ�����顣"
 case 7
 sjpic="Images/img_fy/sj7.gif"
  sjstr="����������������ദ�Ĳ��Ǻܺã�"
 case 8
 sjpic="Images/img_fy/sj8.gif"
  sjstr="������������ĺܺã�ÿ�쳪�����衣"
 case 9
 sjpic="Images/img_fy/sj9.gif"
  sjstr="�������������ķ��˿��Ƿǳ���ʢ��"
 case 10
 sjpic="Images/img_fy/sj10.gif"
  sjstr="�������ڼ����ڼ�������ã����ٿ��ܸ������̡�"
 case 11
 sjpic="Images/img_fy/sj11.gif"
  sjstr="�������ر��ܳԣ��Ѿ��������ǰ���Ŀ��������������������뷨�ټ������ڣ��Ա�������֧��"
 case else
 sjpic="Images/img_fy/sj1.gif"
  sjstr="����������ʹ�����飬����Լ���������Ϊ��"
end select
response.write "<img src="&sjpic&" width=350 height=345 >"
%>
    </td>
  </tr>
  <tr><td align=center><%=sjstr%>
</td></tr>
  <tr><td align=center>���ԣ�<input name=tjly size=30 maxlength=20 > ���250�ַ�</td></tr></center>   
 </table>
<div align="center">��<a href=z_fy_looktj.asp?name=<%=fanren%> title='����鿴����'>����<%=fanren%>�����Ա�</a>��<br><br>
<% if master or jymaster then %>
<input type="submit" value='����Ա����'  name="submit">&nbsp;&nbsp;<input type="button" value="��������" onClick="window.close()" name="button">
<%else%>
<input type="submit" value='��<%=tj_set(1)%>Ԫ����'  name="submit">&nbsp;&nbsp;<input type="button" value="��������" onClick="window.close()" name="button">
<%end if%>
</form>
<%if qj_set(0)="1" then%>
<form action=z_fy_tanjian.asp?action=qiang&name=<%=fanren%> method=post>
&nbsp;<input type="submit" value="�ٺ٣���Ҳ�н��죬����û������" name="submit2">
</form>
<%end if%>
</div>
</body>   
</html>   
<%
end sub

sub savetj()
  if isnull(request.form("tjly")) or request.form("tjly")="" then
   content="ĬĬ���������ᣬ��������������......"
  else
   content=checkstr(request.form("tjly"))
  end if
  mysign="2|0"
  call logs("����","̽��",membername,fanren)
  if (not master) and (not jymaster) then
  sql="update [user] set userwealth=userwealth-"&tj_set(1)&" where username='"&membername&"'"
  conn.execute sql
  end if
  %>
  <html>
<head>
<title><%=membername%>̽������<%=fanren%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body background="Images/img_fy/bg001.gif" >
<table border=0  cellspacing=1 cellpadding=3 align="center" valign="middle" width="60%" ><center>
  <tr> 
    <td align=center  ><p><br><br><%=membername%>̽�ಢ����:<%=content%><br><br><%=fanren%>���ʵ�˵��~~Ҫ��˵ʲô���أ�<br><br>��������һ�������㣡</td></tr></table>
<div align="center"><input type="button" value="�رմ���" onClick="window.close()" name="button">
</div>
</body>   
</html> 
<%  
end sub

sub qiang()
dim qjjg,userw,usere,userc,qjtype,mess1,qjs,qds
if qj_set(4)="1" then
  if session("qjhave")<>"" then
	if session("qjhave")="haveqj" then
	  Errmsg=Errmsg+"<br>"+"<li>��Ҫ̫̰�ģ�20������ֻ����һ��Ŷ��"
         call dvbbs_error()
         exit sub
	end if
  end if
end if
session("qjhave")="haveqj"
  if qj_set(0)="0" then
   mess1="�����������ھ���ɭ�ϣ���ʲôҲû������"
   qds="0"
  else
   sql="select userwealth,userep,usercp from [user] where username='"&fanren&"'"
   set rs=conn.execute(sql)
   userw=rs(0)
   usere=rs(1)
   userc=rs(2)
   randomize 
   qjjg=Int((30*rnd)+1)
   select case qjjg
    case 1
     mess1="ССϺ��ඣ������������ֽ� "
     qjtype="userwealth"
     qjs=0.01
     qds=clng(userw*qjs)
    case 8
     mess1="��ʤ���ޣ���������������ֵ "
     qjtype="userep"
     qjs=0.01
     qds=clng(usere*qjs)
    case 15
     mess1="û�׷���������������ֵ "
     qjtype="usercp"
     qjs=0.01
     qds=clng(userc*qjs)
    case 4
     mess1="�����������ֽ� "
     qjtype="userwealth"
     qjs=0.05
     qds=clng(userw*qjs)
    case 18
     mess1="����ѽ����������ֵ "
     qjtype="userep"
     qjs=0.05
     qds=clng(usere*qjs)
    case 29
     mess1="�����˴����ˣ�Ҫ����������������ֵ "
     qjtype="usercp"
     qjs=0.05
     qds=clng(userc*qjs)
    case 5
     mess1="�����ˣ��Ǻǣ���Ȼ�����ֽ� "
     qjtype="userwealth"
     qjs=0.1
     qds=clng(userw*qjs)
    case 26
     mess1="Ү��������������ֵ "
     qjtype="userep"
     qjs=0.1
     qds=clng(usere*qjs)
    case 23
     mess1="������Ѫǿ�ˣ���������ֵ "
     qjtype="usercp"
     qjs=0.1
     qds=clng(userc*qjs)
    case 10
     mess1="���Ǹ��գ��ٺ٣������ֽ� "
     qjtype="userwealth"
     qjs=0.2
     qds=clng(userw*qjs)
    case 21
     mess1="�����治����������ֵ "
     qjtype="userep"
     qjs=0.2
     qds=clng(usere*qjs)
    case 7
     mess1="Ҫ�ε��ˣ�������������ֵ "
     qjtype="usercp"
     qjs=0.2
     qds=clng(userc*qjs)
    case 11
     mess1="���ģ��Ҿ�Ȼ�Ϸ������ֽ� "
     qjtype="userwealth"
     qjs=0.3
     qds=clng(userw*qjs)
    case else
     mess1="�ţ���С��Ǯ�����Ķ��ˣ�ʲôҲû����!"
     qjtype="none"
     qjs=1
     qds=0
  end select
  if qjtype<>"none" then
     sql="update [user] set "&qjtype&"="&qjtype&"-round("&qds&") where username='"&fanren&"'"  
     conn.execute sql
     sql="update [user] set "&qjtype&"="&qjtype&"+round("&qds&") where username='"&membername&"'"
     conn.execute sql
     if qj_set(1)="1" then
        sql="update [user] set usercp=usercp-"&qj_set(2)&" where username='"&membername&"'"
        conn.execute sql
     end if
     select case qjtype
      case "userwealth"
           qjtype="�ֽ�"
      case "userep"
           qjtype="����"
      case "usercp"
           qjtype="����"
     end select
     content="�������У�"&fanren&"��ǿ������ "&qjtype&":"&qds
     mysign="3|0"
     call logs("����","����",membername,fanren)
  end if
end if
  
  %>
  <script>
confirm('<%=mess1%><%=qds%>')
</script>

  <html>
<head>
<title><%=membername%>����<%=fanren%>�ˣ�<%=fanren%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

</head>
<body background="Images/img_fy/bg001.gif" >
<table border=0  cellspacing=1 cellpadding=3 align="center" valign="middle" width="60%" ><center>
  <tr> 
    <td align=center  ><p><br><br><%=mess1%><%if qds<>"0" then response.write(qds)%><br><br><%if clng(userw)*clng(qjs)<50 or clng(usere)*clng(qjs)<10 or clng(userc)*clng(qjs)<10 then
            response.write " "&fanren&"������������������Ҿ�ʣ����ͷ�ˣ�����~~<br><br>"
    else%>
<%=fanren%>���ʵ�˵���˵�ù����ʲô�¶�������<br><br><%end if%>��������һ���������㣡</td></tr></table>
<div align="center"><input type="button" value="�رմ���" onClick="window.close()" name="button">
</div>
</body>   
</html> 
<%  
end sub

sub checkjail()
if tj_set(0)="0" then
    Errmsg=Errmsg+"<br>"+"<li>�����������ڹر���̽�๦�ܡ�"
    founderr=true
    call dvbbs_error()
    exit sub
end if
fanren=checkstr(request("name"))
sql="select userclass from [user] where username='"&fanren&"' and lockuser=1"
set rs=conn.execute(sql)
if rs.eof then
  Errmsg=Errmsg+"<br>"+"<li>�޴�������������ѳ�����"
  founderr=true
  call dvbbs_error()
  exit sub
end if
sql="select userwealth from [user] where username='"&membername&"'"
set rs=conn.execute(sql)
if rs(0)<clng(tj_set(1)) and not master and not jymaster then
  Errmsg=Errmsg+"<br>"+"<li>����̽�ӷѲ�����"
  founderr=true
  call dvbbs_error()
  exit sub
end if
end sub
%>
