<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
if aqjh_name<>Application("aqjh_user") then Response.Redirect "../error.asp?id=486"
%>
<html><head>
<script>
function copyText(obj,b_cmd) {var rng = document.body.createTextRange();rng.moveToElementText(obj);rng.select();rng.execCommand(b_cmd);}
</script>
<title>账号急救</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head><LINK href=css/css.css type=text/css rel=stylesheet>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center">**************<b></b><b>账号急救</b>**************<b></b></p>
<p><font size="2" color="#FF0000">当玩家账号发生意外或被误斩<br>
  可使用此功能从备份数据库中恢复账号<br>
备份的江湖数据库位于江湖目录/DataBackup<br>
  </font><br>
</p>
<table width="50%" border="2" align="center" cellpadding="10" cellspacing="10">
  <tr>
    <td><p align="center"><strong>账号急救</strong> </p>
      <form name="form1" method="post" action="RUSER2.ASP">
        <p><font size="2">请输入玩家的名字：</font><br>
          <br>
          <input type="text" name="getname" size="20" maxlength="150">
          <br>
          <font size="2"><br>
          请输入选择数据库的名字<br>
          <br>
<%
if IsObjInstalled("Scripting.FileSystemObject") then
Set Fso=server.createobject("scripting.filesystemobject")
dim method
method=request.querystring("method")
FoldName=server.mappath("../DataBackup") 
Set f1 = Fso.GetFolder(FoldName)
Set pFiles = f1.Files
%>
<select name="DBpath" onclick="javascript:document.form1.getmdb.value=this.value;">
<option value="">请选择</option>
<%for each file in f1.Files
tempwrite= tempwrite & "<option value=" & file.name &">" & file.name &"</option>"
next 
response.write tempwrite
set f1=nothing
set fso=nothing
%>
</select>
<%
Set f1 = Nothing
Set pFiles = Nothing
Set Fso=Nothing
else
%>
FSO权限已被您的服务器管理员关闭！
<%end if%>
<input type="text" name="getmdb" size="20" maxlength="150" value="数据库名.asp">
        </p>        <p align="center">
          <input type="submit" name="ggc" value="·执 行·">
          <br>
        </p>
      </form>
</td></tr></table></body></html>
<%
Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function
%>