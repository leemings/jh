<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
pai=session("yx8_mhjh_usercorp")
sql="Select * from 门派 where 门派='"&pai&"'"
rs=conn.execute(sql)
if pai="无" then
%>
<script language=vbscript>
MsgBox "错误！你还没有加入任何门派，请回吧！"
location.href = "javascript:history.back()"
</script>
<%else%>
<html>
<head>
<title>本门密室</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../style.css">
</head>
<body bgcolor="#000000" text="#FFFFFF" oncontextmenu=self.event.returnValue=false>
<p align="center"><b><font color="#FFFF00" size="4">本门密室</font></b></p>
<p align="center">
<marquee style="color: #FFFF00">本门公告:<%=RS("公告")%></marquee>
</p>
<form action="mpjianjie1.asp?subject=<%=pai%>" method=POST>
<ul>
<table width="438" border="0" cellspacing="0" cellpadding="0" align="CENTER" height="145">
<tr>
<td height="145" width="603">
<table width="436" border="0" cellspacing="0" cellpadding="0" height="223">
<tr>
<td width="4" height="223"></td>
<td  width="413" align="center" height="223">
<div align="center" style="width: 413; height: 222">
<table border=1 cellspacing=0 cellpadding=3 align="center" width="330" bordercolordark="#FFFFFF" bordercolorlight="#cccc99">
<tr>
<td align="center" width="62" valign="middle" height="25">
<font class="c" color="#FFFFFF">门&nbsp;&nbsp;&nbsp; 派</font></td>         
<td width="250"> <font color="#FFFFFF"> <%=pai%> </font> </td>       
</tr>       
<tr>       
<td align="center" width="62" valign="middle" height="25">       
<font color="#FFFFFF">       
公&nbsp;&nbsp;&nbsp; 告</font></td>         
<td width="250">       
<font color="#FFFFFF">       
<input name="comment" value="<%=RS("公告")%>" size=30 maxlength=50>(50字内）</font>       
</td>       
</tr>       
<tr>       
<td align="center" width="62" valign="middle" height="25">       
<font color="#FFFFFF">       
简&nbsp;&nbsp;&nbsp; 介</font></td>         
<td width="250"><font color="#FFFFFF"><input name="nomment" value="<%=RS("简介")%>" size=30 maxlength=50>(50字内）</font> </td>       
</tr>       
<tr>       
<td align="center" width="62" valign="middle" height="25">       
<font class="c" color="#FFFFFF">掌&nbsp;&nbsp;&nbsp; 门</font></td>           
<td width="250">          
<font color="#FFFFFF">       
<%       
set conn=server.createobject("adodb.connection")       
conn.open Application("yx8_mhjh_connstr")       
set rst=server.CreateObject ("adodb.recordset")       
sql="select 姓名 from 用户 where 门派='"&pai&"' and 身份='掌门'"       
set rst=conn.execute(sql)       
if rst.EOF or rst.BOF then       
%>       
空缺       
<%else%>       
<%=rst("姓名")%>       
<%end if%>	                 </font>	                 </td>       
</tr>       
<tr>       
<td align="center" width="62" valign="middle"><font color="#FFFFFF">加入条件</font></td>       
<td width="250"><font color="#FFFFFF"><%if RS("加入条件")="False"  then%>       
拒绝任何人加入       
<%elseif RS("加入条件")="True" then%>       
欢迎所有加入       
<%else%>       
<%=RS("加入条件")%>       
<%end if%>       
</font>       
</td>       
</tr>       
<tr>       
<td align="center" colspan="2" valign="middle" height="46" width="320">       
<font color="#FFFFFF">       
<input type="HIDDEN" name="action" value="RegSubmit">       
<input type="SUBMIT" name="Submit" value="更新">       
<input type="reset" name="Submit2" value="清空">         
&nbsp;</font>         
</td>       
</tr>       
<tr>       
<td align="center" colspan="2" valign="middle" height="44" width="320"><font color="#FFFFFF">帮主可以在这里更改本帮的一些简要介绍</font></td>       
</tr>       
</table>       
</div>       
</td>       
<td width="13" height="223" ></td>       
</tr>       
</table>       
</td>       
</tr>       
</table>       
</ul>       
</form>       
</body>       
</html>       
<%       
end if       
conn.close    
set conn=nothing       
%>       
       
