<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=210"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	%>
	<script language="vbscript">
	 alert "你不能进行操作，进行此操作必须进入聊天室！"
	 window.close()
	</script>
	<%
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * from 用户 where 姓名='"&sjjh_name&"'",conn
if rs("身份")<>"掌门" then
	%>
	<script language="vbscript">
	 alert "你又不是掌门，乱跑什么！"
	 window.close()
	</script>
	<%
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
randomize timer
regjm=int(rnd*9998)+1
%>
<html>
<head>
<title>〖<%=Application("sjjh_chatroomname")%>〗</title>
<style type="text/css">body, td     { font-size: 14 }
input        { font-size: 14; color: #000000 }
.p1          { font-size: 21pt; color: #ff0000 }
.p2          { font-size: 9pt; color: #00ee00 }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body topmargin="0" background="../JHIMG/Bk_hc3w01.gif">
<p class="p1" align="center"><b>掌门禅位</b></p>
<table border="1" bgcolor="#FFFFFF" align="center" width="428" cellpadding="10"
cellspacing="13">
  <tr>
    <td background="../bgcheetah.gif"> 
      <form method="POST" action="rangweiok.asp">
<table width="100%">
<tr>
            <td><font size="-1">下任掌门</font><font size="-1">:      
              <input type="text" name="xrzm" size="10" maxlength="10">    
              </font> </td>    
            <td><font size="-1">你的密码:      
              <input type="password" name="password" size="10" maxlength="10">    
              </font> </td>    
</tr>    
<tr>    
            <td height="2"><font size="-1">认证:      
              <input type=text name=regjm1 size=5 maxlength="5"> 
              </font> </td>                
            <td height="2"> <font size="-1"><br>    
              防止朔雪输入认证:<font color="#FF0000"><%=regjm%></font></font></td>    
</tr>    
<tr>    
<td align="center" colspan="2"><input type=hidden name=regjm value="<%=regjm%>"> 
              <input type="submit" value="掌门禅位"    
name="submit"> 
               <input type="reset" value="取消" name="reset"></td>    
</tr>    
<tr>    
            <td align="center" colspan="2"><b><font color="#FF0000" size="2">掌门禅位：对方战斗等级必须在50级以上</font></b></td>    
</tr>    
</table>    
</form>    
</td>    
</tr>    
</table>    
    
</body>    
    
</html> 
 
