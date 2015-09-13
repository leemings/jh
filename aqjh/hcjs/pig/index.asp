<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT id FROM 用户 WHERE  姓名='" & aqjh_name &"'",conn,2,2
aqjh_id=rs("id")
rs.close
%>
<!--#include file="data.asp"-->
<%
rs.open "Select * from 猪圈设置",connt,3,3
ks=rs("投票开始")
jz=rs("投票结束")
dj=rs("等级")
rs.close
rs.open "Select * from 养猪投票者 where 投票ID="&aqjh_id,connt,3,3
if rs.eof or rs.bof then
	tp=0
else
	tp=1
end if
rs.close
sql="Select count(*) As 数量 from 养猪投票者"
set tmprs=connt.execute(sql)
ytrs=tmprs("数量")
tmprs.close
set tmprs=nothing
sql="Select count(*) As 数量 from 养猪风云榜"
set tmprs=connt.execute(sql)
hxr=tmprs("数量")
tmprs.close
set tmprs=nothing
sql="select * from 养猪风云榜 order by 票数 desc"
set rs=connt.execute(sql)
%>
<HTML><HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<TITLE>〖养猪大赛〗</TITLE>
<link rel="stylesheet" href="../css.css">
</HEAD>
<BODY bgcolor="#FFFFFF" background="../../jhimg/bk_hc3w.gif" topmargin=0>
<!--#include file="head.asp"-->
<table width=500 border="0" align=center cellspacing="0" bordercolor="#FF0000">
        <tr bordercolor="#FFFFFF">
          <td align=center><font color=red><b>爱 情 江 湖 养 猪 大 赛 投 票 处</b></font><BR><BR></td>
        </tr>
      </table>
      <table width=650 border="1" align=center cellspacing="0" bordercolorlight="#800080"
bordercolordark="#FFFFFF">
        <tr bgcolor="#DFEFFF">
          <td width=90 align=center>参赛人</td>
          <td width=90 align=center>参赛主题</td>
          <td width=270 align=center>主题说明</td>
          <td width=100 align=center>颁奖(等级)</td>
          <td width=60 align=center>得票率</td>
          <td width=40 align=center>得票</td><td width=80 align=center>投我一票</td>
<!--#include file="data.asp"-->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
sql="select * from 养猪风云榜 order by 票数 desc"
rs.open sql,connt,1,1
if rs.recordcount<>0 then
  rs.pagesize = 25
  page = cint(request("pageno"))
     if page <= 1 then 
        page = 1
     end if
     if page >= rs.pagecount then
        page = rs.pagecount
     end if
     rs.absolutepage = page
  for i=1 to rs.pagesize
%>
        <tr bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';">
          <td width=90 align=center><font color=red><%=rs("参赛人")%></font></td>
          <td width=90 align=center><a href=list.asp?myid=<%=rs("参赛人")%> target=_blank><font color=blue><%=rs("参赛主题")%></font></a></td>
          <td width=270 align=center><%=rs("主题说明")%></td>
          <td width=100 align=center><%if aqjh_grade>=10 then%><a href=jpset.asp?name=<%=rs("参赛人")%>&j=1 target=_blank>[1]</a> <a href=jpset.asp?name=<%=rs("参赛人")%>&j=2 target=_blank>[2]</a> <a href=jpset.asp?name=<%=rs("参赛人")%>&j=3 target=_blank>[3]</a> <a href=jpset.asp?name=<%=rs("参赛人")%>&j=0 target=_blank>[清]</a><%else%>无权操作<%end if%></td>
          <td width=60 align=center>
<%if ytrs=0 then%>
        <div align="center">0%            
          <%else
	n=int(rs("票数")/ytrs*100)
	%>
          <%=n%>%            
          <%end if%>
</td>
<td width=40 align=center><font color=red><%=rs("票数")%></font></td><td align=center>
<%if CDate(ks)>now() then%>
投票未开始
<%elseif CDate(jz)<now() then%>
投票已结束
<%elseif aqjh_jhdj<dj then%>
等级不够
<%elseif tp=1 then%>
已经投过票
<%else%>
<a href=tp.asp?name=<%=rs("参赛人")%> target=_blank>投我一票</a>
<%end if%>
</td></tr>
<% 
  rs.movenext
  if rs.eof then exit for
  next
 %>   
共<font color=2020dd><% =rs.pagecount %></font>页<font color=2020dd><% =rs.recordcount %></font>张 当前:<font color=2020dd><%=page%></font>/<%=rs.pagecount%>&nbsp;&nbsp;<a href=index.asp?pageno=1>首页</a>&nbsp;<a href=index.asp?pageno=<% =page-1 %>>上页</a>&nbsp;<a href=index.asp?pageno=<% =page+1 %>>下页</a>&nbsp;<a href=index.asp?pageno=<% =rs.pagecount %>>末页</a>
<%end if%>
</table><br>
<center><FONT color=#0000ff>&copy; 版权所有 2015-2015 </FONT><A href="http://www.happyjh.com/" target=_blank><FONT color=#0000ff>快乐江湖网</FONT></A>
</body></html>