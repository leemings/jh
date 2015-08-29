<% if Session("myName")="" then
  response.write("SORRY！您没有<a href=../index.htm>登录</a>系统。")
else%>
<!--#include file="data1.asp"--><%
sql="SELECT * FROM 房屋 where  类型='一般房屋'"
Set Rs=conn.Execute(sql)
name=session("myname")
%>
<html>
<LINK 
href="../pic/css.css" rel=stylesheet>
<body bgcolor="#FFFFFF">
<table border=1 bgcolor="#FFFFFF" align=center width=54% cellpadding="0" cellspacing="1" bordercolor="#ff7010">
  <tr bgcolor="#006633"> 
    <td height="30" colspan="7" bgcolor="#FFFFFF"> 
      <p align=center class="p9"><font style="FONT-SIZE: 9pt" color="#FF3333"><marquee><font color="#990099">欢迎</font><%=name%><font color="#990099">光临房产中心，谁都想有自己的家，这里可以满足您的各种要求：(现在卖出房子只能得到原来的1/5的价格)</font></marquee></font></p>
  <tr bgcolor="#006633"> 
    <td height="30" colspan="7" bgcolor="#FFFFFF"> 
      <table width="100%" border="1" cellspacing="1" cellpadding="1" bordercolor="#6633FF">
        <tr> 
          <td bgcolor="#9999FF"> 
            <div align="center"><a href=fangwu.asp><font color="#FF3333">一般房屋</font></a></div>
          </td>
          <td bgcolor="#9999FF"> 
            <div align="center"><a href=fangwu3.asp><font color="#FF3333">高级公寓</font></a></div>
          </td>
          <td bgcolor="#9999FF"> 
            <div align="center"><a href=fangwu4.asp><font color="#FF3333">花园洋房</font></a></div>
          </td>
          <td bgcolor="#9999FF"> 
            <div align="center"><a href=fangwu5.asp><font color="#FF3333">豪华别墅</font></a></div>
          </td>
        </tr>
      </table>
  <tr bgcolor=#74E76D bordercolor="#000000"> 
    <td height="15" width="14%" bgcolor="#FF9933" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">房屋类型</font></div>
    </td>
    <td height="15" width="9%" bgcolor="#FF9933" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">售价</font></div>
    </td>
    <td height="15" width="11%" bgcolor="#FF9933" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF3333">城市</font></div>
    </td>
    <td height="15" colspan="2" bgcolor="#FF9933" bordercolor="#FF6600">
      <div align="center"><font size="2" color="#FF3333">门牌号码</font></div>
    </td>
    <td height="15" width="19%" bgcolor="#FF9933" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF3333">房主/伴侣</font></div>
    </td>
    <td height="15" width="25%" bgcolor="#FF9933" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">操作</font></div>
    </td>
  </tr>
  <% 
do while not rs.eof and not rs.bof 
%> 
  <tr bgcolor=#DEAD63> 
    <td width="14%" bgcolor="#FFFFFF" class="calen-curr" height="15"><font size="2"> 
      <center>
        <%=rs("类型")%> 
      </center>
      </font> 
    <td width="9%" bgcolor="#FFFFFF" class="calen-curr" height="15"><%=rs("售价")%> 
      <div align="center"></div>
    <td width="11%" bgcolor="#FFFFFF" class="ct-def4" height="15"> 
      <div align="center"><font size="2"><%=rs("位置")%></font></div>
    </td>
    <td colspan="2" bgcolor="#FFFFFF" class="ct-def4" height="15"><%=rs("街道")%> 
      <font color=red><%=rs("id")%></font><font size="2">号</font></td>
    <td width="19%" bgcolor="#FFFFFF" class="calen-curr" height="15"> 
      <div align="center"><font size="2"><%=rs("户主")%> / </font><font color="ff00ff"><%=rs("伴侣")%></font></div>
    </td>
    <td width="25%" bgcolor="#FFFFFF" class="calen-curr" height="15"><font size="2"> 
      <center>
        <%if rs("户主")<>name and rs("户主")="无" then%><a href=fangwu1.asp?id=<%=rs("id")%>><font color="0000ff">购买</font></a><%end if%><%if rs("户主")=sjjh_name then%><a href=fangwu2.asp?id=<%=rs("id")%>><font color=red>卖出</font></a><%end if%><%if rs("户主")<>"无" and rs("户主")<>name then%>该房己出售<%end if%> 
      </center>
      </font></td>
  </tr>
  <% 
rs.movenext 
loop 
conn.close 
%> 
</table>
</html>
<%end if%>