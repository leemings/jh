<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<TITLE>〖养猪大赛〗</TITLE>
<link rel="stylesheet" href="../../css.css">
</HEAD>
<BODY bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif" topmargin=0>
<!--#include file="head.asp"-->
<table width=500 border="0" align=center cellspacing="0" bordercolor="#FF0000">
        <tr bordercolor="#FFFFFF">
          <td align=center><font color=red><b>爱 情 江 湖 养 猪 大 赛 风 云 榜</b></font><BR><BR></td>
        </tr>
      </table>
<!--#include file="data.asp"-->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.open "Select * from 猪圈设置",connt,3,3
jz=rs("投票结束")
if cdate(jz)>now() then
 Response.Write "<center>提示：投票没结束，尚不能开奖！"
 Response.End
end if
%>
      <table width=650 border="1" align=center cellspacing="0" bordercolorlight="#800080"
bordercolordark="#FFFFFF">
        <tr bgcolor="#DFEFFF">
          <td width=90 align=center>参赛人</td>
          <td width=90 align=center>参赛主题</td>
          <td width=270 align=center>主 题 说 明</td>
          <td width=100 align=center>猪赛日期</td>
          <td width=40 align=center>名次</td><td width=40 align=center>奖 品</td>
<%
rs.close
sql="select * from 养猪风云榜 where 名次>0 order by 名次"
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
          <td width=90 align=center><a href=list.asp?myid=<%=rs("参赛人")%> target=_blank><font color=blue><%=rs("参赛主题")%></a></td>
          <td width=270 align=center><%=rs("主题说明")%></td>
          <td width=100 align=center><%=rs("猪赛日期")%></td>
          <td width=40 align=center>第<font color=red><%=rs("名次")%></font>名</td><td width=40 align=center><%if rs("领奖")=true then%>已领奖<%else%><a href=hfybok.asp?myid=<%=rs("参赛人")%> target=_blank>领奖</a><%end if%></td>
        </tr>
<% 
  rs.movenext
  if rs.eof then exit for
  next
 %>   
共<font color=2020dd><% =rs.pagecount %></font>页<font color=2020dd><% =rs.recordcount %></font>张 当前:<font color=2020dd><%=page%></font>/<%=rs.pagecount%>&nbsp;&nbsp;<a href=hfyb.asp?pageno=1>首页</a>&nbsp;<a href=hfyb.asp?pageno=<% =page-1 %>>上页</a>&nbsp;<a href=hfyb.asp?pageno=<% =page+1 %>>下页</a>&nbsp;<a href=hfyb.asp?pageno=<% =rs.pagecount %>>末页</a>
<%end if%>
</table><br>
<center>
<font color=red>注意：领奖前请先进行杀猪，否则系统将自动清理你的猪崽，以备下一届养猪大赛的举行！</font><br><br>
<FONT color=#0000ff>&copy; 版权所有 2015-2015 </FONT><A href="http://www.happyjh.com/" target=_blank><FONT color=#0000ff>快乐江湖网</FONT></A>
</body></html>