<!--#include file="conn.asp"-->
<!--#include file="z_fy_Conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_fy_const.asp"-->
<!--#include file="z_plus_check.asp"-->
<%
'=========================================================
' Plusname: ��Ժ����������
' File: z_fy_fayuan.asp
' For Version: 6.0sp2
' Date: 2003-2-10
' Script Written by Hero20000
'=========================================================
stats="������Ժ"
call nav()
call head_var(2,0,"","")
if  founderr then
   call dvbbs_error()
else
   call getfyconfig()
   response.write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1 ><tr><th valign=middle colspan=2 align=center height=25><b>��Ժ����</b></td></tr><tr><td valign=middle class=tablebody2 height=100><CENTER>"
   response.write "<table border=0 width=97% cellspacing=0 cellpadding=0 ><tr><td width=40% align=center class=TopLighNav1>"
   if master then
   response.write "<a href=z_fy_manage.asp title='���������������̨��������'><IMG height=22 border=0 src=Images/img_fy/logo1.gif width=22 align=absMiddle>���٣�</a>"
   else
   response.write "<IMG height=22 border=0 src=Images/img_fy/logo1.gif width=22 align=absMiddle>���٣�"
   end if
   if not isnull(fyname) and  fyname<>"" and  fyname<>"��" then response.write "<a href=dispuser.asp?name="&fyname&" target=_blank>"&fyname&"</a>(<font color=red>����</font>) "
   sql="select username from [user] where usergroupid=1"    
   set rs=conn.execute(sql)    
   do while Not RS.Eof 
        response.write "<a href=dispuser.asp?name="&rs(0)&" target=_blank>"&checkstr(rs(0))&"</a> "	
        rs.movenext
   loop     
  dim fyall,fyover
  set fyall=connfy.execute("select count(id) from [fy]")
  set fyover=connfy.execute("select count(id) from [fy] where result<>'N'") 
  response.write "</td><td width=20% align=center class=TopLighNav1><IMG height=22 src=Images/img_fy/logo2.gif width=22 align=absMiddle><a href=z_fy_write.asp>���ϴ�</a></td>"
  response.write "<td width=20% align=center class=TopLighNav1><IMG height=22 src=Images/img_fy/logo5.gif width=22 align=absMiddle><a href=z_fy_over.asp>״ֽ�嵥</a></td>" 
  response.write "<td width=20% align=center class=TopLighNav1><IMG height=22 src=Images/img_fy/logo9.gif width=22 align=absMiddle><a href=z_fy_fanren.asp>�ιۼ���</a></td></tr></table>"
  response.write "<marquee scrollamount=3 onmouseover=this.stop(); onmouseout=this.start();><img src=Images/img_fy/gg.gif width=16 height=15><b>������ԺĿǰ������<font color=blue> "&fyall(0)&" </font>��Ͷ�ߣ������<font color=blue> "&fyover(0)&" </font>��������<font color=red> "&fyall(0)-fyover(0)&"</font> ����������</b></marquee>"
  response.write "<br><table  width=97% cellspacing=1 cellpadding=3 class=tableborder2><tr><td rowspan=2 width=37% class=tablebody2 valign=top><br>"&memo_set&"<br>"
  response.write "<br><img src=Images/img_fy/fy.gif></td><td rowspan=2  width=3% class=tablebody2> </td><td></td></tr>"
  response.write "<tr><td width=60% class=tablebody2><table width=100% align=center cellspacing=1 cellpadding=3 bgcolor="&Forum_Body(27)&"><tr><td width=100% class=tablebody1><br>"
  sql= "SELECT top 1 * FROM fy ORDER BY dateandtime DESC"  
  set rs=connfy.execute(sql)
  if not rs.eof and not rs.bof then              
     response.write "<div align=center><b>���ֵ�<font color=red>"&rs(0)&"</font>��</b></div><br><hr><br>"     
     response.write "<b>ԭ��</b>��"&rs(4)&"<br><br>"
     response.write "<b>����</b>��"&rs(2)&"<br><br>"
     response.write "<b>����/����ʱ��</b>��"&rs(8)&"<br><br>"
     response.write "<b>����</b>��"&rs(1)&"<br><br>"
     response.write "<b>���龭����</b>"&rs(3)&"<br><br>"
     response.write "<b>ԭ��Ҫ��</b><font color=red>"&rs(5)&"</font><br><br>"  
     response.write "<b>�������ߣ�</b>"&rs(9)&"<br>"
     if lcase(membername)=lcase(rs(2)) and rs(7)="N" then response.write "<div align=right><a href=z_fy_shensu.asp?id="&rs(0)&"><font color=red><b>�����л�Ҫ˵>>></b></font></a></div>"
     response.write "<hr><b>���ٳ´ʣ�</b><font color=red>"&rs(10)&"</font><br><br>"
     response.write "<b>�о������</b>"
     select case  rs(7)
       case "N" 
         response.write "���޽��<br><br>"
       case "B"
         response.write "ԭ�����<br><br>"
       case "P"
         response.write "��ƽ��<br><br>"
       case "C"
         response.write "ԭ���ѳ���<br><br>"
       case else
        response.write "�ۺ�������������豻��<b><font color=#0000ff>"&rs(7)&"</font></b>�Ĵ�����<br><br>"
     end select
     response.write "<b>Ŀǰ״̬��</b>"
     select case rs(6)
       case "N"  
         response.write "�ȴ�������" 
       case "0"
           response.write "��ʵ��������������"
           if fjbh_set(0)<>"0" and fjbh_set(0)<>"+0" and fjbh_set(0)<>"-0" then response.write "��ԭ���ֽ���"&fjbh_set(0)&" "
           if fjbh_set(1)<>"0" and fjbh_set(1)<>"+0" and fjbh_set(1)<>"-0" then response.write "�������ֽ���"&fjbh_set(1)&" " 
        case "1" 
           response.write "��ʵȷ�䣬��ԭ��Ҫ��Ա���ִ���˴���"
           if fjty_set(0)<>"0" and fjty_set(0)<>"+0" and fjty_set(0)<>"-0" then response.write "��ԭ���ֽ���"&fjty_set(0)&" "
           if fjty_set(1)<>"0" and fjty_set(1)<>"+0" and fjty_set(1)<>"-0" then response.write "�������ֽ���"&fjty_set(1)&" "
        case "2"
          response.write "�����飬��ԩ�ٴ�����ƽ��"
          if fjpf_set(0)<>"0" and fjpf_set(0)<>"+0" and fjpf_set(0)<>"-0" then response.write "��ԭ���ֽ���"&fjpf_set(0)&" "
          if fjpf_set(1)<>"0" and fjpf_set(1)<>"+0" and fjpf_set(1)<>"-0" then response.write "�������ֽ���"&fjpf_set(1)&" "
        case "4" 
          response.write "ԭ���ܸ棬�������"
          if fjzy_set(0)<>"0" and fjzy_set(0)<>"+0" and fjzy_set(0)<>"-0" then response.write "��ԭ���ֽ���"&fjzy_set(0)&" "
          if fjzy_set(1)<>"0" and fjzy_set(1)<>"+0" and fjzy_set(1)<>"-0" then response.write "�������ֽ���"&fjzy_set(1)&" "
        case "5" 
          response.write "�����ܸ棬�ط�ԭ��"
          if fjwg_set(0)<>"0" and fjwg_set(0)<>"+0" and fjwg_set(0)<>"-0" then response.write "��ԭ���ֽ���"&fjwg_set(0)&" "
          if fjwg_set(1)<>"0" and fjwg_set(1)<>"+0" and fjwg_set(1)<>"-0" then response.write "�������ֽ���"&fjwg_set(1)&" "
        case "6"
          response.write "ԭ�����߲�����ʵ�������������о�" 
        case "7"
          response.write "���������������ѱ��ͷ�"
        case "8"
          response.write "ԭ���ѳ���"
      end select    
   else
     response.write "<div align=center><b>�ΰ�����,û���˱�����</b>!</div><br>"
    end if
  rs.close  
  set rs=nothing 
  fyall.close
  fyover.close
  set fyall=nothing
  set fyover=nothing
  response.write "</td></tr></table></td></tr></table><br></CENTER></td></tr></table>"  
end if  
call activeonline()
call footer()  
%>   
