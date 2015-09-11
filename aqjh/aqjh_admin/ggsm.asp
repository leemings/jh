<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<!--#include file="config.asp"-->
<%
cz=trim(request.querystring("cz"))
if cz<>"备份" then cz=""
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
%>
<html>
<head>
<title>聊天室顶部广告修改</title><LINK href="css/css.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center">聊天室顶部广告修改(&lt;br&gt;为回车 支持html语言)
<p><br>
如果资料修改错了,会造成一些想不到的后果，如果进入不了聊天室等，如果你修改错了请点<a href="ggsm.asp?cz=%B1%B8%B7%DD"><b><font color="#FF0000">这里</font></b></a>进行必要的修正，再点修改就可以恢复到初始化的状态 
  ，注意修改不要错了!</font>
<%
rs.open "SELECT * FROM sm where a='进入"&cz&"'",conn,2,2%>
<form method=POST action="ggsmok.asp?cz=进入" name="form">
  <div align="center"><font color="#0000FF"><b>进入江湖的广告设置</b></font><br>
    (非会员，每一条之间以;分号半角分隔)<br>
    <textarea name="hysm"  cols="90" rows="8"><%=rs("c")%></textarea>
    <br>
    (会员进入说明，只有一句话\n表示换行!) <br>
    <textarea name="hydz"  cols="90" rows="4"><%=rs("d")%></textarea>
    <br>
    <br>
    <br>
    <input type=submit value=确定进入广告修改 name="submit" style="border: 1px solid; font-size: 9pt; border-color:#000000 solid">
  </div>
</form>
<%rs.close
rs.open "SELECT * FROM sm where a='pk值日'",conn,2,2%>
<form method=POST action="ggsmok.asp?cz=pk值日" name="form">
  <div align="center"><font color="#0000FF"><b>官府值日表</b></font><br>
    <textarea name="hysm"  cols="90" rows="8"><%=rs("c")%></textarea>
    <br>
    <input type=submit value=确定修改值日表 name="submit" style="border: 1px solid; font-size: 9pt; border-color:#000000 solid">
  </div>
</form>
<%rs.close
rs.open "SELECT * FROM sm where a='滚动"&cz&"'",conn,2,2%>
<form method=POST action="ggsmok.asp?cz=滚动" name="form">
  <div align="center"><font color="#0000FF"><b>聊天室上面滚动广告设置</b></font><br>
    每一条之间以;分号半角分隔<br>
    <textarea name="hysm"  cols="90" rows="12"><%=rs("c")%></textarea>
    <br>
    江湖博士格式如下：问1|答案1;问2|答案2;<br>
    <textarea name="hydz"  cols="90" rows="12"><%=rs("d")%></textarea>
    <br>
    <br>
    <br>
    <input type=submit value=确定滚动广告修改 name="submit" style="border: 1px solid; font-size: 9pt; border-color:#000000 solid">
  </div>
</form>
<%rs.close
set rs=nothing
conn.close
set conn=nothing%> </body>
</html>