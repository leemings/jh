<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
house=Request("house")
for each element in Request.Form
elevalue=Request.Form(element)
if instr(elevalue,"'")<>0 or instr(elevalue,chr(34))<>0 or instr(elevalue,"\")<>0 or instr(elevalue,";") or instr(elevalue,"?") or instr(elevalue,"or") or instr(elevalue,"&") or instr(elevalue,"#")then Response.Redirect "../error.asp?id=056"
next
name=session("yx8_mhjh_username")
if name="" then
response.redirect "../error.asp?id=016"
response.end
end if
if house=empty or house="" then
 house="一般房屋"
elseif house="b" then
 house="高级公寓"
elseif house="c" then
 house="花园洋房"
elseif house="d" then
 house="豪华别墅"
else
 house="一般房屋"
end if
%>
<!--#include file="data1.asp"-->
<%
sql="SELECT * FROM 房屋 where 类型='"&house&"'"
Set Rs=conn.Execute(sql)
%>
<html>
<head>
<title>房 产 中 心</title>
<style type="text/css">
<!--
td {  font-family: "宋体"; font-size: 12px}
-->
</style>
</head>
<body bgcolor=#000000 text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF">
<table align=center border=0 cellpadding=0 cellspacing=0 width=608 background="table_bg.gif">
  <tbody> 
  <tr> 
    <td background="history_top_bg.gif" valign=top width=608 height="7"> 
      <table align=center border=0 cellpadding=2 cellspacing=0 class=p9 
      width="100%">
        <tbody> 
        <tr> 
          <td height=25 width="31%">&nbsp;</td>
          <td height=25 valign=top width="30%"> 
            <div align=center><font color="#000000">房产中心</font></div>
          </td>
          <td height=25 valign=top 
width="39%">&nbsp;</td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
  <tr> 
    <td valign=top width="608"> 
      <div align="center"><br>
        房 产 中 心 欢迎<%=name%>光临房产中心，谁都想有自己的家，这里可以满足您的各种要求：        
        <br>       
        <br>       
      </div>       
      <table width='95%' align=center cellspacing=0 border=0 cellpadding=0>       
        <tr bgcolor="#993300">        
          <td>        
            <center>       
              <a href=fangwu.asp>一般房屋</a>        
            </center>       
          </td>       
          <td>        
            <center>       
              <a href=fangwu.asp?house=b>高级公寓</a>        
            </center>       
          </td>       
          <td>        
            <center>       
              <a href=fangwu.asp?house=c>花园洋房</a>        
            </center>       
          </td>       
          <td height="20">        
            <center>       
              <a href=fangwu.asp?house=d>豪华别墅</a>        
            </center>       
          </td>       
        </tr>       
      </table>       
      <table width='95%' align=center cellspacing=0 border=1 cellpadding=0 bordercolor="#003333">       
        <tr>        
          <td align="center" bgcolor="#CC6699">房屋类型</td>       
          <td align="center" bgcolor="#CC6699">售价</td>       
          <td align="center" bgcolor="#CC6699">城市</td>       
          <td align="center" bgcolor="#CC6699">门牌号码</td>       
          <td align="center" bgcolor="#CC6699">拥有人/伴侣</td>       
          <td align="center" height="20" bgcolor="#CC6699">操作</td>       
        </tr>       
<%        
do while not rs.eof and not rs.bof        
%>        
        <tr bgcolor="#6666FF">        
          <td align="center"><img src=<%=rs("图片")%> width="30" height="30"><%=rs("类型")%></td>       
          <td align="center"><%=rs("售价")%></td>       
          <td align="center"><%=rs("位置")%></td>       
          <td align="center"><%=rs("街道")%><font color=#FFFF99><%=rs("id")%></font>号</td>       
          <td align="center"><%=rs("户主")%><img src=image/heart.gif width="15" height="15"><%=rs("伴侣")%></td>       
          <td align="center" height="24"><%if rs("户主")<>name and rs("户主")="无" then%><a href=fangwu1.asp?id=<%=rs("id")%>>购买</a><%end if%><%if rs("户主")=name then%><a href=fangwu2.asp?id=<%=rs("id")%>>卖出</a><%end if%><%if rs("户主")<>"无" and rs("户主")<>name then%>该房己出售<%end if%></td>       
        </tr>       
<%        
rs.movenext        
loop   
rs.Close                
set rs=nothing                
conn.Close                
set conn=nothing               
%>        
      </table>           
    </td>       
  </tr>       
  <tr>        
    <td bgcolor=#847939 height=1 width="608"> </td>       
  </tr>       
  </tbody>        
</table>       
</body>        
</html> 