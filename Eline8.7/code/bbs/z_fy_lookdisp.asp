<!--#include file="conn.asp"-->
<!--#include file="z_fy_Conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_fy_const.asp"-->
<%
'=========================================================
' Plusname: 法院监狱第三版
' File: z_fy_lookdisp.asp
' For Version: 6.0sp2
' Date: 2003-2-10
' Script Written by Hero20000
'=========================================================
dim id,fyrs,fysql
id=request.querystring("id")
call getfyconfig()
fysql="SELECT * FROM fy WHERE ID="& id
set fyrs=connfy.execute(fysql)
%>
<html>
<head>
<title>查看状纸详情</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body >
<!-- #include file="inc/FORUM_CSS.asp" -->    
<table border=0  cellspacing=1 cellpadding=3 align="center" width="100%" class=tableborder1>
  <tr> 
    <td  height="24" class=tablebody2>原  告 <a href=dispuser.asp?name=<%=fyrs(4)%> target=_blank><b><%=fyrs(4)%></b></a>&nbsp;愤怒地写道： 
    </td>
  </tr>
  <tr>
    <td class=tablebody1><br>
      <%=changechr(fyrs(3))%><br>
      <br>
      希望给于 <b><font color="#FF0000"><%=fyrs(5)%></font></b> 的处罚。
     </td>
   </tr>
  <tr> 
    <td height="24" class=tablebody2>被  告 <a href=dispuser.asp?name=<%=fyrs(2)%> target=_blank><b><%=fyrs(2)%></b></a>&nbsp;申诉道：<br>
    </td>
  </tr>
  <tr>
    <td class=tablebody1><br>
      <%=changechr(fyrs(9))%>
      <br><br>
    </td>
  </tr>  
 <tr> 
    <td  height="24" class=tablebody2>法官判决如下:  
    </td>
  </tr>
  <tr>
    <td class=tablebody1>
      <%=changechr(fyrs(10))%><br>
      <br>
<%
select case  fyrs(7)
       case "N" 
         response.write "尚无结果<br><br>"
       case "B"
         response.write "原告败诉<br><br>"
       case "P"
         response.write "已平反<br><br>"
       case "C"
         response.write "原告已撤诉<br><br>"
       case else
        response.write "综合上述情况，给予被告<b><font color=#0000ff>"&fyrs(7)&"</font></b>的处罚。<br><br>"
     end select
%>
     </td>
   </tr>
   <tr> 
    <td  height="24" class=tablebody2>目前状态:  
    </td>
  </tr>
  <tr><td class=tablebody1>
<% select case fyrs(6)
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
fyrs.close
set fyrs=nothing
%></td></tr>   
 </table>
 
<br>
<div align="center"><input type="button" value="关闭" onClick="window.close()" name="button">
</div>
</body>   
</html>   
<%   
function changechr(str)   
    changechr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," ","&nbsp;")   
    changechr=replace(replace(replace(replace(changechr,"[img]","<img src="),"[b]","<b>"),"[red]","<font color=CC0000>"),"[big]","<font size=7>")   
    changechr=replace(replace(replace(replace(changechr,"[/img]","></img>"),"[/b]","</b>"),"[/red]","</font>"),"[/big]","</font>")   
end function   
%>