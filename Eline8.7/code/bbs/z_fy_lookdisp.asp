<!--#include file="conn.asp"-->
<!--#include file="z_fy_Conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_fy_const.asp"-->
<%
'=========================================================
' Plusname: ��Ժ����������
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
<title>�鿴״ֽ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body >
<!-- #include file="inc/FORUM_CSS.asp" -->    
<table border=0  cellspacing=1 cellpadding=3 align="center" width="100%" class=tableborder1>
  <tr> 
    <td  height="24" class=tablebody2>ԭ  �� <a href=dispuser.asp?name=<%=fyrs(4)%> target=_blank><b><%=fyrs(4)%></b></a>&nbsp;��ŭ��д���� 
    </td>
  </tr>
  <tr>
    <td class=tablebody1><br>
      <%=changechr(fyrs(3))%><br>
      <br>
      ϣ������ <b><font color="#FF0000"><%=fyrs(5)%></font></b> �Ĵ�����
     </td>
   </tr>
  <tr> 
    <td height="24" class=tablebody2>��  �� <a href=dispuser.asp?name=<%=fyrs(2)%> target=_blank><b><%=fyrs(2)%></b></a>&nbsp;���ߵ���<br>
    </td>
  </tr>
  <tr>
    <td class=tablebody1><br>
      <%=changechr(fyrs(9))%>
      <br><br>
    </td>
  </tr>  
 <tr> 
    <td  height="24" class=tablebody2>�����о�����:  
    </td>
  </tr>
  <tr>
    <td class=tablebody1>
      <%=changechr(fyrs(10))%><br>
      <br>
<%
select case  fyrs(7)
       case "N" 
         response.write "���޽��<br><br>"
       case "B"
         response.write "ԭ�����<br><br>"
       case "P"
         response.write "��ƽ��<br><br>"
       case "C"
         response.write "ԭ���ѳ���<br><br>"
       case else
        response.write "�ۺ�������������豻��<b><font color=#0000ff>"&fyrs(7)&"</font></b>�Ĵ�����<br><br>"
     end select
%>
     </td>
   </tr>
   <tr> 
    <td  height="24" class=tablebody2>Ŀǰ״̬:  
    </td>
  </tr>
  <tr><td class=tablebody1>
<% select case fyrs(6)
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
fyrs.close
set fyrs=nothing
%></td></tr>   
 </table>
 
<br>
<div align="center"><input type="button" value="�ر�" onClick="window.close()" name="button">
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