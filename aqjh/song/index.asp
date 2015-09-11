<!-- #include file="conn.asp"-->
<%
aqjh_grade=session("aqjh_grade")
if session("aqjh_name")=""  then Response.Redirect "../../error.asp?id=440"
%>
<html>
<head>
<LINK href=STYLE.CSS type=text/css rel=stylesheet>
<title><%=Application("aqjh_chatroomname")%> - 点歌列表</title>
<%

sql="select * from music order by dateandtime desc"

%>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<table width="99%"align="center" cellPadding="0"  border="1" bgcolor="#F78B4A" bordercolor="#FF3300" bordercolorlight="#FFFFFF"cellspacing="0" >
  <tr>
<td bgcolor="#FDA782" colspan="6">
<%
mypage=request.QueryString("whichpage")
if mypage="" then
mypage=1
end if 
mypagesize=request.QueryString("pagesize")
if mypagesize="" then 
mypagesize=10
end if 
set rs = Server.CreateObject("adodb.recordset")
rs.cursorlocation=aduseclient
rs.cachesize=8
rs.open sql,conn,1,1
rs.movefirst
rs.pagesize=mypagesize
maxcount=cint(rs.pagecount)
rs.absolutepage=mypage
howmanyfields=rs.fields.count-1
tal=rs.RecordCount
response.Write"第"&mypage&"页，总共（"&maxcount&"）页总共有（"&tal&"）份祝福<br>"
%>
</td>
</tr>
<tr align="center">
<td bgcolor="#FF6666">点歌人</td>
<td bgcolor="#FF6666">赠给</td>
<td bgcolor="#FF6666">祝福</td>
<td bgcolor="#FF6666">歌曲</td>
<td bgcolor="#FF6666">日期</td>
<%if aqjh_grade>7 then%>
<td bgcolor="#FF6666">操作</td>
<%end if%>
</tr>
<%
tn=0
do while not rs.eof and tn < rs.pagesize %>
<tr align="center">
<td bgcolor="#FFFEEE"><%=rs("name")%></td>
<td bgcolor="#FFFEEE"><%=rs("toname")%></td>
<td bgcolor="#FFFEEE">〈<a href="#" onClick="MM_openBrWindow('look.asp?id=<%=rs("id")%>','zhufu','width=300,height=120')">查看祝福</a>〉</td>
<td bgcolor="#FFFEEE">〈<a href="#" onClick="MM_openBrWindow('Play.asp?id=<%=rs("id")%>','url','width=250,height=80')"><%=rs("songname")%></a>〉</td>
<td bgcolor="#FFFEEE"><%=left(rs("dateandtime"),len(cstr(rs("dateandtime")))-8)%></td>
<%if aqjh_grade>7 then%>
<td bgcolor="#FFFEEE"><a href="#" onClick="window.open('xg.asp?id=<%=rs("id")%>','xg','width=330,height=190')">修改</a>&nbsp;&nbsp;&nbsp;<a href="del.asp?id=<%=rs("id")%>">删除</a></td>
<%end if%>
</tr>
<tr>
<td  bgcolor="#FDA782" colspan="6">
<%
rs.MoveNext
tn=tn+1
loop
%>
<%
pad="0"
scriptname=request.ServerVariables("SCRIPT_NAME")
for counter =1 to maxcount
if counter>=10 then
pad=""
end if 
ref="<a href='"&scriptname&"?whichpage="&counter
ref=ref&"'>"&pad&counter&"页""</a>"
response.write ref&""
if counter mod  10 =  0 then
response.Write"<br>"
end if 
next
%>
</td>
</tr>
</table>