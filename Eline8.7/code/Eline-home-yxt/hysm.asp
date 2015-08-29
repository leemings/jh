<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
if sjjh_name<>Application("sjjh_user") then 
	Response.Write "<script Language=Javascript>alert('提示：你不是正站长，你不能操作！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
cz=trim(request.querystring("cz"))
if cz<>"备份" then cz=""
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM sm where a='会员"&cz&"'",conn,2,2
%>
<html>
<head>
<title>会员说明及联系方式修改</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif">
<p align="center">会员说明及联系方式修改(&lt;br&gt;为回车 支持html语言)<br>
  (<font color="#0000FF">办理会员说明</font>) <br>
  <input type=submit value=恢复初始设置 onClick="location.href ='hysm.asp?cz=备份'" name="submit2" style="border: 1px solid; font-size: 9pt; border-color:#000000 solid">
<form method=POST action="hysmok.asp" name="form">
  <div align="center">
    <textarea name="hysm"  cols="60" rows="14"><%=rs("c")%></textarea>
    <br>
    <br>
    (<font color="#0000FF">联系方法说明</font>) <br>
    <textarea name="hydz"  cols="60" rows="5"><%=rs("d")%></textarea>
    <br>
    <input type=submit value=确定会员修改 name="submit" style="border: 1px solid; font-size: 9pt; border-color:#000000 solid">
  </div>
</form>
    </body>
</html>
