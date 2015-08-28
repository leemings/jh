
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="select 杀人 from 用户 where 姓名='" & username & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
response.redirect "../error.asp?id=016"
conn.close
response.end
else
sm=rs("杀人")
rs.close
conn.close
%>

<head>
<link rel="stylesheet" href="../style.css">
<title>地狱</title>
</head>

<body bgcolor="#000000" text="#FFFFFF" oncontextmenu=self.event.returnValue=false>
<table width="493" border="0" cellspacing="0" cellpadding="0" align="CENTER" height="267">
<tr>
<td height="267" width="508">
<table width="508" border="0" cellspacing="0" cellpadding="0" height="255">
<tr>
<td height="255" width="1" ></td>
<td  background="../image/haidian_0928_2.jpg" width="504" align="center" height="255"><br>
<div id="Layer2" style="z-index: 2; left: 208; width: 267; position: absolute; top: 283; height: 80">
<table cellspacing="0" border="0" width="251" bordercolorlight="#cccc99" bordercolordark="#FFFFFF" height="15">
<% if sm<=5 then
response.write "你想做通缉犯？你现在还不够格，先去杀点人吧！"
response.end
else%>
<form method=POST action=diyu2.asp  onSubmit="this.ok.disabled=true">
<tr>
<td height="1" width="247">
<div align="left">你的杀人数字为<%=sm%>, <%  
if sm<=1 then response.write "不需要炼了！再炼就会又变成通缉犯了！":p=1  
if sm<11 and sm>0 then response.write "实行一级炼火！":p=2  
if sm<=20 and sm>10 then response.write "实行二级炼火":p=3  
if sm>=21 then response.write "江湖重犯，实行三味真火！":p=4  
%> </div>  
<tr>  
<td width="247">  
<div align="left">修炼方法：  
<input name=yilao readonly size=12 value="<%  
if p=1 then response.write "回去吧,不欢迎你"  
if p=2 then response.write "一级炼火"  
if p=3 then response.write "二级炼火"  
if p=4 then response.write "三味真火"  
%>">  
<input name=ok type=submit value=确定>  
&nbsp;  
</div>  
</td>  
</tr>  
</form>  
<%end if%>  
</table>  
</div>  
</td>  
<td height="255" width="1">  
</div>  
</td>  
</tr>  
</table>  
</td>  
</tr>  
</table>  
</body>  
<%end if  
rs.close  
set rs=nothing  
conn.close  
set conn=nothing%>  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  


