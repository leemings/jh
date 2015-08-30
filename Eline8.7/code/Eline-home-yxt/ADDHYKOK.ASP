<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../pass.asp"-->
<%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
knum=int(abs(Request.form("knum")))
kpass=LCase(Request.form("kpass"))
kauto=LCase(request.form("kauto"))
if kauto="on" then 
kpass=""
for i=1 to 8 step 1
randomize timer
kpass=kpass+cstr(int(rnd*10))
next
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_hykmdb")
rs.open "SELECT * FROM sethyk Where name='set'",conn
kname="HY0"+cstr(rs("js"))
rs.close
rs.open "SELECT * FROM 会员卡 Where 类型=4",conn
if rs.eof or rs.bof then
conn.execute "insert into 会员卡(卡号,密码,点数,类型,用户ID,时间) values ('"&kname&"','"&md5(kpass)&"',"&knum*1000&",0,0,0)"
else
conn.execute "update 会员卡 set 卡号='"&kname&"',密码='"&md5(kpass)&"',点数="&knum*1000&" ,类型=0,用户ID=0,时间=0 where 卡号='"&rs("卡号")&"'"
end if
rs.close
conn.execute "update sethyk set js=js+1 where name='set'"
conn.close
set rs=nothing
set conn=nothing
%>
<html>
<head>
<title>会员数据库管理♀wWw.happyjh.com♀

</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>
<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif">
<p align="center"><font color="#0000FF">会员卡管理</font></p>
<p align="center">　</p>
<p align="center">　</p>
<p align="center">会员卡已经建立完毕</p>
<table border="1" width="24%" bordercolorlight="#000000" cellspacing="1" cellpadding="1" bordercolordark="#85C2E0" height="79" align="center">
  <tr> 
    <td bgcolor="#85C2E0"   width="51%" align="center" nowrap height="30">会 员 卡ID：</td> 
 
    <td bgcolor="#85C2E0"   width="51%" align="center" nowrap height="30"><%=kname%></td> 
 </tr> 
  <tr>  
    <td bgcolor="#85C2E0"   width="51%" align="center" nowrap height="31">会员卡密码：</td> 
    <td bgcolor="#85C2E0"   width="51%" align="center" nowrap height="31"><%=kpass%></td> 
  </tr> 
  <tr>  
    <td bgcolor="#85C2E0"   width="51%" align="center" nowrap height="31">会员卡金额：</td> 
    <td bgcolor="#85C2E0"   width="51%" align="center" nowrap height="31"><%=cstr(int(abs(Request.form("knum"))))%></td> 
  </tr> 
</table> 
<p align="center"><font color="#FF0000">请注意，密码经过了MD5数据库加密，所以请牢记密码，否则将没有办法查阅！</font></p> 
<p align="center">休闲江湖制作</p>
</body>
</html>