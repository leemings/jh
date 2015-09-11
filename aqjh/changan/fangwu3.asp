<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
id=request("id")
my=aqjh_name
%>
<!--#include file="data1.asp"--><%
sql="SELECT * FROM 房屋 where  类型='高级公寓'"
Set Rs=conn.Execute(sql)
%>
<html>
<LINK 
href="../css.css" rel=stylesheet>
<body bgcolor="#FFFDDF">
<table border=1 bgcolor="#588878" align=center width=98% cellpadding="0" cellspacing="1" bordercolor="#65251C">
  <tr bgcolor="#006633"> 
    <td height="15" colspan="6" bgcolor="#588878"> 
      <p align="center"> <img src="aqwjy_pic.jpg" width="311" height="93" align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <img src="aqwjy_zi.gif" width="126" height="93"> 
</table>
<table border=1 bgcolor="#FFFFCC" align=center width=98% cellpadding="0" cellspacing="1" bordercolor="#65251C">
  <tr> 
    <td height="30" colspan="7" bgcolor="#FFFFFF"> 
      <p align=center class="p9"><font style="FONT-SIZE: 9pt" color="#FF3333"><marquee><font color="#990099">欢迎</font><%=name%><font color="#990099">光临房产中心，谁都想有自己的家，这里可以满足您的各种要求：(现在卖出房子只能得到原来的1/5的价格)</font></marquee></font></p>
  <tr> 
    <td height="21" colspan="7" bgcolor="#FFFFFF"> 
      <table width="100%" border="1" cellspacing="1" cellpadding="1" bordercolor="#6633FF">
        <tr> 
          <td bgcolor="#65251C"> 
            <div align="center"><a href=fangwu.asp><font color="#FFFFFF">一般房屋</font></a></div>
          </td>
          <td bgcolor="#65251C"> 
            <div align="center"><a href=fangwu3.asp><font color="#FFFFFF">高级公寓</font></a></div>
          </td>
          <td bgcolor="#65251C"> 
            <div align="center"><a href=fangwu4.asp><font color="#FFFFFF">花园洋房</font></a></div>
          </td>
          <td bgcolor="#65251C"> 
            <div align="center"><a href=fangwu5.asp><font color="#FFFFFF">豪华别墅</font></a></div>
          </td>
        </tr>
      </table>
  <tr bordercolor="#000000"> 
    <td height="15" width="14%" bgcolor="#65251C" bordercolor="#65251C"> 
      <div align="center"><font size="2" color="#FFFFFF">房屋类型</font></div>
    </td>
    <td height="15" width="9%" bgcolor="#65251C" bordercolor="#65251C"> 
      <div align="center"><font size="2" color="#FFFFFF">售价</font></div>
    </td>
    <td height="15" width="11%" bgcolor="#65251C" bordercolor="#65251C"> 
      <div align="center"><font size="2" color="#FFFFFF">城市</font></div>
    </td>
    <td height="15" colspan="2" bgcolor="#65251C" bordercolor="#65251C"> 
      <div align="center"><font size="2" color="#FFFFFF">门牌号码</font></div>
    </td>
    <td height="15" width="19%" bgcolor="#65251C" bordercolor="#65251C"> 
      <div align="center"><font size="2" color="#FFFFFF">房主/伴侣</font></div>
    </td>
    <td height="15" width="25%" bgcolor="#65251C" bordercolor="#65251C"> 
      <div align="center"><font size="2" color="#FFFFFF">操作</font></div>
    </td>
  </tr>
  <% 
do while not rs.eof and not rs.bof 
%> 
  <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="14%" class="calen-curr" height="15"><font size="2"> 
      <center>
        <%=rs("类型")%> 
      </center>
      </font> 
    <td width="9%" class="calen-curr" height="15"><%=rs("售价")%> 
      <div align="center"></div>
    <td width="11%" class="ct-def4" height="15"> 
      <div align="center"><font size="2"><%=rs("位置")%></font></div>
    </td>
    <td colspan="2" class="ct-def4" height="15"><%=rs("街道")%> <font color=red><%=rs("id")%></font><font size="2">号</font></td>
    <td width="19%" class="calen-curr" height="15"> 
      <div align="center"><font size="2"><%=rs("户主")%> / </font><font color="ff00ff"><%=rs("伴侣")%></font></div>
    </td>
    <td width="25%" class="calen-curr" height="15"><font size="2"> 
      <center>
        <%if rs("户主")<>name and rs("户主")="无" then%><a href=fangwu1.asp?id=<%=rs("id")%>><font color="0000ff">购买</font></a><%end if%><%if rs("户主")=aqjh_name then%><a href=fangwu2.asp?id=<%=rs("id")%>><font color=red>卖出</font></a><%end if%><%if rs("户主")<>"无" and rs("户主")<>name then%>该房己出售<%end if%> 
      </center>
      </font></td>
  </tr>
  <% 
rs.movenext 
loop 
set rs=nothing	
	conn.close
	set conn=nothing
%> 
</table>
</html>
<html><script language="JavaScript">                                                                  </script></html>

