<!--#include file="conn.asp"-->
<!--#include file="z_fy_Conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_fy_const.asp"-->
<!--#include file="z_plus_check.asp"-->
<%
'=========================================================
' Plusname: 法院监狱第三版
' File: z_fy_fayuan.asp
' For Version: 6.0sp2
' Date: 2003-2-10
' Script Written by Hero20000
'=========================================================
stats="社区法院"
call nav()
call head_var(2,0,"","")
if  founderr then
   call dvbbs_error()
else
   call getfyconfig()
   response.write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1 ><tr><th valign=middle colspan=2 align=center height=25><b>法院公告</b></td></tr><tr><td valign=middle class=tablebody2 height=100><CENTER>"
   response.write "<table border=0 width=97% cellspacing=0 cellpadding=0 ><tr><td width=40% align=center class=TopLighNav1>"
   if master then
   response.write "<a href=z_fy_manage.asp title='法官请点这里进入后台参数设置'><IMG height=22 border=0 src=Images/img_fy/logo1.gif width=22 align=absMiddle>法官：</a>"
   else
   response.write "<IMG height=22 border=0 src=Images/img_fy/logo1.gif width=22 align=absMiddle>法官："
   end if
   if not isnull(fyname) and  fyname<>"" and  fyname<>"无" then response.write "<a href=dispuser.asp?name="&fyname&" target=_blank>"&fyname&"</a>(<font color=red>特邀</font>) "
   sql="select username from [user] where usergroupid=1"    
   set rs=conn.execute(sql)    
   do while Not RS.Eof 
        response.write "<a href=dispuser.asp?name="&rs(0)&" target=_blank>"&checkstr(rs(0))&"</a> "	
        rs.movenext
   loop     
  dim fyall,fyover
  set fyall=connfy.execute("select count(id) from [fy]")
  set fyover=connfy.execute("select count(id) from [fy] where result<>'N'") 
  response.write "</td><td width=20% align=center class=TopLighNav1><IMG height=22 src=Images/img_fy/logo2.gif width=22 align=absMiddle><a href=z_fy_write.asp>诉讼处</a></td>"
  response.write "<td width=20% align=center class=TopLighNav1><IMG height=22 src=Images/img_fy/logo5.gif width=22 align=absMiddle><a href=z_fy_over.asp>状纸清单</a></td>" 
  response.write "<td width=20% align=center class=TopLighNav1><IMG height=22 src=Images/img_fy/logo9.gif width=22 align=absMiddle><a href=z_fy_fanren.asp>参观监狱</a></td></tr></table>"
  response.write "<marquee scrollamount=3 onmouseover=this.stop(); onmouseout=this.start();><img src=Images/img_fy/gg.gif width=16 height=15><b>社区法院目前共受理<font color=blue> "&fyall(0)&" </font>条投诉，已审结<font color=blue> "&fyover(0)&" </font>条，尚有<font color=red> "&fyall(0)-fyover(0)&"</font> 条正在审理</b></marquee>"
  response.write "<br><table  width=97% cellspacing=1 cellpadding=3 class=tableborder2><tr><td rowspan=2 width=37% class=tablebody2 valign=top><br>"&memo_set&"<br>"
  response.write "<br><img src=Images/img_fy/fy.gif></td><td rowspan=2  width=3% class=tablebody2> </td><td></td></tr>"
  response.write "<tr><td width=60% class=tablebody2><table width=100% align=center cellspacing=1 cellpadding=3 bgcolor="&Forum_Body(27)&"><tr><td width=100% class=tablebody1><br>"
  sql= "SELECT top 1 * FROM fy ORDER BY dateandtime DESC"  
  set rs=connfy.execute(sql)
  if not rs.eof and not rs.bof then              
     response.write "<div align=center><b>法字第<font color=red>"&rs(0)&"</font>号</b></div><br><hr><br>"     
     response.write "<b>原告</b>："&rs(4)&"<br><br>"
     response.write "<b>被告</b>："&rs(2)&"<br><br>"
     response.write "<b>受理/处理时间</b>："&rs(8)&"<br><br>"
     response.write "<b>理由</b>："&rs(1)&"<br><br>"
     response.write "<b>案情经过：</b>"&rs(3)&"<br><br>"
     response.write "<b>原告要求：</b><font color=red>"&rs(5)&"</font><br><br>"  
     response.write "<b>被告申诉：</b>"&rs(9)&"<br>"
     if lcase(membername)=lcase(rs(2)) and rs(7)="N" then response.write "<div align=right><a href=z_fy_shensu.asp?id="&rs(0)&"><font color=red><b>被告有话要说>>></b></font></a></div>"
     response.write "<hr><b>法官陈词：</b><font color=red>"&rs(10)&"</font><br><br>"
     response.write "<b>判决结果：</b>"
     select case  rs(7)
       case "N" 
         response.write "尚无结果<br><br>"
       case "B"
         response.write "原告败诉<br><br>"
       case "P"
         response.write "已平反<br><br>"
       case "C"
         response.write "原告已撤诉<br><br>"
       case else
        response.write "综合上述情况，给予被告<b><font color=#0000ff>"&rs(7)&"</font></b>的处罚。<br><br>"
     end select
     response.write "<b>目前状态：</b>"
     select case rs(6)
       case "N"  
         response.write "等待审理中" 
       case "0"
           response.write "事实不符，驳回起诉"
           if fjbh_set(0)<>"0" and fjbh_set(0)<>"+0" and fjbh_set(0)<>"-0" then response.write "，原告现金另"&fjbh_set(0)&" "
           if fjbh_set(1)<>"0" and fjbh_set(1)<>"+0" and fjbh_set(1)<>"-0" then response.write "，被告现金另"&fjbh_set(1)&" " 
        case "1" 
           response.write "事实确凿，按原告要求对被告执行了处罚"
           if fjty_set(0)<>"0" and fjty_set(0)<>"+0" and fjty_set(0)<>"-0" then response.write "，原告现金另"&fjty_set(0)&" "
           if fjty_set(1)<>"0" and fjty_set(1)<>"+0" and fjty_set(1)<>"-0" then response.write "，被告现金另"&fjty_set(1)&" "
        case "2"
          response.write "经调查，属冤假错案，已平反"
          if fjpf_set(0)<>"0" and fjpf_set(0)<>"+0" and fjpf_set(0)<>"-0" then response.write "，原告现金另"&fjpf_set(0)&" "
          if fjpf_set(1)<>"0" and fjpf_set(1)<>"+0" and fjpf_set(1)<>"-0" then response.write "，被告现金另"&fjpf_set(1)&" "
        case "4" 
          response.write "原告诬告，提出警告"
          if fjzy_set(0)<>"0" and fjzy_set(0)<>"+0" and fjzy_set(0)<>"-0" then response.write "，原告现金另"&fjzy_set(0)&" "
          if fjzy_set(1)<>"0" and fjzy_set(1)<>"+0" and fjzy_set(1)<>"-0" then response.write "，被告现金另"&fjzy_set(1)&" "
        case "5" 
          response.write "严重诬告，重罚原告"
          if fjwg_set(0)<>"0" and fjwg_set(0)<>"+0" and fjwg_set(0)<>"-0" then response.write "，原告现金另"&fjwg_set(0)&" "
          if fjwg_set(1)<>"0" and fjwg_set(1)<>"+0" and fjwg_set(1)<>"-0" then response.write "，被告现金另"&fjwg_set(1)&" "
        case "6"
          response.write "原告所诉部分属实，法官已酌情判决" 
        case "7"
          response.write "被告刑期已满，已被释放"
        case "8"
          response.write "原告已撤诉"
      end select    
   else
     response.write "<div align=center><b>治安良好,没有人被起诉</b>!</div><br>"
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
