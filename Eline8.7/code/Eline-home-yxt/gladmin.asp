<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>
<title>��E�߽���������Ա���ó����wWw.51eline.com��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>
<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif">
<div align="center"><font size="+2" face="����_GB2312">��E�߽���������Ա���ó���</font><br>
  <%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM �û� where ����='�ٸ�' or grade>=6 order by -grade",conn,2,2
do while not rs.bof and not rs.eof%>
</div>
<form method=POST action='gladminok.asp'>
<div align="center">
<input type="text" name="name"  readonly size="15" maxlength="10" value="<%=rs("����")%>">
<select name=grade size=1  style="BACKGROUND-COLOR: #99CCFF; BORDER-BOTTOM: 1px double; BORDER-LEFT: 1px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000">
 <option value="6" <%if rs("grade")=6 then%>selected<%end if%>>6</option>
 <option value="7" <%if rs("grade")=7 then%>selected<%end if%>>7</option>
 <option value="8" <%if rs("grade")=8 then%>selected<%end if%>>8</option>
 <option value="9" <%if rs("grade")=9 then%>selected<%end if%>>9</option>
 <option value="10" <%if rs("grade")=10 then%>selected<%end if%>>10</option>   
 </select>
 <input type="text" name="mp" readonly size="15" maxlength="10" value="<%=rs("����")%>">
 <input type="text" name="sf"  size="15" maxlength="10" value="<%=rs("���")%>">
<input type=submit value=�޸� name="submit" style="border: 1px solid; font-size: 9pt; border-color:#000000 solid">
<input type=submit value=ɾ�� name="submit" style="border: 1px solid; font-size: 9pt; border-color:#000000 solid">
</div>
</form>
<%
rs.movenext
loop
conn.close
%>
<form method=POST action='gladminok.asp'>
<div align="center">
<input type="text" name="name"   size="15" maxlength="10" value="">
<select name=grade size=1  style="BACKGROUND-COLOR: #99CCFF; BORDER-BOTTOM: 1px double; BORDER-LEFT: 1px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000">
 <option value="6" selected>6</option>
 <option value="7" >7</option>
 <option value="8" >8</option>
 <option value="9" >9</option>
 <option value="10" >10</option>   
 </select>
<input type=submit value=��� name="submit" style="border: 1px solid; font-size: 9pt; border-color:#000000 solid">
</div>
</form>
</html>