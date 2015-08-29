<%PageName="SongConAdd"%>
<!--#include file="function.asp"-->
<%CheckAdmin1%>
<!--#include file="conn.asp"-->
<%
Specialid=request.QueryString("Specialid")
set rs=server.createobject("adodb.recordset")
i=0
sql="select * from Special where Specialid="&cstr(Specialid)
rs.open sql,conn,1,3
if rs.EOF then
	errmsg= "<li>服务器出错！请联系管理员</li>"
	call error()
	Response.End
else
i=i+1
%>
<html> 
<head> 
<title>Untitled Document</title> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"> 
<style type="text/css"> 
<!-- 
td { font-size: 9pt} 
a { color: #000000; text-decoration: none} 
a:hover { text-decoration: underline} 
.tx { height: 16px; width: 30px; border-color: black black #000000; border-top-width: 0px; border-right-width: 0px; border-bottom-width: 1px; border-left-width: 0px; font-size: 9pt; background-color: #FFD2E4; color: #0000FF} 
.bt { font-size: 9pt; border-top-width: 0px; border-right-width: 0px; border-bottom-width: 0px; border-left-width: 0px; height: 16px; width: 80px; background-color: #FFD2E4; cursor: hand} 
.tx1 { height: 20px; width: 30px; font-size: 9pt; border: 1px solid; border-color: black black #000000; color: #0000FF} 
INPUT{BORDER-TOP-WIDTH: 1px; PADDING-RIGHT: 1px; PADDING-LEFT: 1px; BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; BORDER-LEFT-COLOR: #cccccc; BORDER-BOTTOM-WIDTH: 1px; BORDER-BOTTOM-COLOR: #cccccc; PADDING-BOTTOM: 1px; BORDER-TOP-COLOR: #cccccc; PADDING-TOP: 1px; HEIGHT: 18px; BORDER-RIGHT-WIDTH: 1px; BORDER-RIGHT-COLOR: #cccccc}
textarea {border-width: 1; border-color: #000000; background-color: #efefef; font-family: 宋体; font-size: 9pt; font-style: bold;}
select {border-width: 1; border-color: #000000; background-color: #efefef; font-family: 宋体; font-size: 9pt; font-style: bold;}
--> 
</style> 
</head> 
<div align="center">
<center>
<table border="1" width="99%" cellspacing="0" cellpadding="3" class="TableLine" bordercolor="#FF1171" bordercolordark="#FFFFFF">
<tr> 
 <td width="100%" height="20" colspan=3 bgcolor="#FF1171" align=center><font color="white"><b>批 量 添 加 歌 曲</b></font></td> 
</tr> 
<tr align="left" valign="middle" bgcolor="#FFD2E4"> 
<td bgcolor="#FFD2E4" height="92">
<form name="form1" method="POST" action="SongConSave.asp"> 
<script language="javascript"> 
function setid() 
{ 
str='<br>'; 
if(!window.form1.upcount.value) 
window.form1.upcount.value=1; 
for(i=1;i<=window.form1.upcount.value;i++) 
str+='歌曲名称'+i+':<input name="MusicName'+i+'" size=20><input type="text" name="Wma'+i+'" value="'+i+'.Wma" size=6><select name="Classic'+i+'" size="1"><option name=option<%=rs("Classid")%><%=i%> selected value="<%=rs("Classid")%>,<%=rs("SClassid")%>,<%=rs("NClassid")%>,<%=rs("Specialid")%>"><%=rs("Name")%></option></select><br><br>'; 
window.upid.innerHTML=str+'<br>'; 
} 
</script> 
<li>需要添加几首 
<input type="text" name="upcount" class="tx" value="1"> 
<input type="button" name="Button" class="bt" onclick="setid();" value="· 设定 ·"> 
</li> 
<br> 
<br> 
<li>公用地址: 
<input type="text" name="url" class="tx" style="width:300" value="/"> 
<font color="red"><b>注意:必须正确填写该项!</b></font></li> 
</td> 
</tr> 
<tr align="center" valign="middle"> 
<td align="left" id="upid" height="122">
</td> 
</tr> 
<tr align="center" valign="middle" bgcolor="#FFD2E4"> 
<td bgcolor="#FFD2E4" height="24"> 
<input type="submit" name="Submit" value=" 提交 ">&nbsp;  
<input type="reset" name="Submit2" value=" 重执 "> 
</td> 
</tr> 
</table> 
</form> 
<script language="javascript"> 
setid(); 
</script>
<%
end if
rs.close
%>
<%
set rs=nothing
conn.close
set conn=nothing
%>


