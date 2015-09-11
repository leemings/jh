<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.charset="gb2312"%>
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"



aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
aqjh_id=Session("aqjh_id")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"



Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
zuzhuangtai=rs("组队")
zuname=rs("组名")


%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<title>组队</title>
<style type="text/css">
<!--
body {  font-size: 9pt; color=000000}
td {  font-size: 9pt; color=000000}
A {text-decoration: none; color: #76f76f; font-family: "宋体"; font-size: 9pt}
a:hover      { color: ffffFF; font-family: 宋体; font-size: 9pt; text-decoration:
underline }
A:active {  font-family: "宋体"; font-style: normal;  color: ffffFF; text-decoration: underline; font-size: 9pt}
br {font-size:6pt; line-height:80%;}
-->
</style>

</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#02C158" bgproperties="fixed" text="#FFFFFF" link="#FFFFFF" topmargin="0" vlink="#FFFFFF" leftmargin="0">
<div align="center">
  <table  border="0" cellspacing="0" cellpadding="2" width="128" bordercolorlight="#000000" bordercolordark="#FFFFFF"  height="486" >
    <tr> 
      <td align="center"> <%if rs("转生")>=2 and rs("道德")>=3000 then%><a href="jlzudui.asp?id=建立组队">建立组</a><% end if %></td>
    </tr>
    <%
rs.close
rs.open "select * FROM 组队 WHERE 组名='"&zuname&"'",conn
if not(rs.eof or rs.bof) then

%>


	  <%if rs("组长")=aqjh_id then%>
       <tr> 
      <td align="center">        
          <a href="jlzudui.asp?id=删除组队">删除组</a></td>
    </tr>
 
    
    

    <tr> 
      <td align="center"> 
<table width='90%' align=center border=1 cellspacing="0" cellpadding="3" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="<%=chatbgcolor%>">
<form method=post action=tjsc.asp id=form1 name=form1>
    <input type=hidden name=newmo value="">
    <tr> 
      <td width="50%">对某人</td>
      <td width="50%"> 
        <input name='money' type=text maxlength=5 size=8>
      </td>
    </tr>
    <tr align="center"> 
      <td colspan="2"> <input type=radio name=op checked value='加组'>加组
       
		<br>
      <input type=radio name=op value='踢出'>踢出组</td>
    </tr>
    <tr align="center"> 
      <td colspan="2"> 
        <input type=submit value='确 认' name="submit">
      </td>
    </tr></form>
</table> 
  
      </td>
    </tr>
<%end if%>
    <% else %>
    
    
        <tr> 
      <td align="center"> 
       
          <a href="yunxu.asp">打开关闭组队状态</a></td>
    </tr>

        <%
    end if
rs.close
rs.open "select 组名 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
if rs("组名")<>"" or rs("组名")<>"无" then
zuname=rs("组名")
    end if
    rs.close
rs.open "select * FROM 组队 WHERE 组名='"&zuname&"'",conn
if not(rs.eof or rs.bof) then

%>
   <tr> 
      <td align="center"> 
     
          <a href="likai.asp">离开组</a></td>
    </tr>
    <tr> 
      <td align="center"> 
     
          本组人员[<%=rs("组员数")%>]人</td>
    </tr>
    <tr> 
      <td align="center"> 
        
         <%=rs("组员")%> 　</td>
    </tr><% else%>
    <tr> 
      <td align="center"> 

        你还没有加入组  　</td>
    </tr>         
         
<%end if
rs.close%>

    <tr>
 <td align="center"> 
       
          组最大人数<%=Application("aqjh_zudui")%></td>
    </tr>
<tr>
 <td align="center"> 
       
          　</td>
    </tr>
<tr>
 <td align="center"> 
        
          　</td>
    </tr>
<tr>
 <td align="center"> 
       
          　</td>
    </tr>
<tr>
 <td align="center"> 
       
          　</td>
    </tr>
<tr>
      <td height="2" align="center"> 
          
       
      </td>
    </tr>
  </table>
  
</div>
</body>
</html>

    <%
			set rs=nothing
			conn.close
			set conn=nothing

%>
</body>
</html>