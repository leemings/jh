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
rs.open "SELECT id FROM �û� WHERE  ����='" & aqjh_name &"'",conn,2,2
aqjh_id=rs("id")
rs.close
%>
<!--#include file="data.asp"-->
<%
rs.open "Select * from ��԰����",connt,3,3
ks=rs("ͶƱ��ʼ")
jz=rs("ͶƱ����")
dj=rs("�ȼ�")
rs.close
rs.open "Select * from ��԰ͶƱ�� where ͶƱID="&aqjh_id,connt,3,3
if rs.eof or rs.bof then
	tp=0
else
	tp=1
end if
rs.close
sql="Select count(*) As ���� from ��԰ͶƱ��"
set tmprs=connt.execute(sql)
ytrs=tmprs("����")
tmprs.close
set tmprs=nothing
sql="Select count(*) As ���� from ��԰���ư�"
set tmprs=connt.execute(sql)
hxr=tmprs("����")
tmprs.close
set tmprs=nothing
sql="select * from ��԰���ư� order by Ʊ�� desc"
set rs=connt.execute(sql)
%>
<HTML><HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<TITLE>����԰������</TITLE>
<link rel="stylesheet" href="../css.css">
</HEAD>
<BODY bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif" topmargin=0>
<!--#include file="head.asp"-->
<table width=500 border="0" align=center cellspacing="0" bordercolor="#FF0000">
        <tr bordercolor="#FFFFFF">
          <td align=center><font color=red><b>��  ��  �� �� �� ԰ �� �� Ͷ Ʊ ��</b></font><BR><BR></td>
        </tr>
      </table>
      <table width=650 border="1" align=center cellspacing="0" bordercolorlight="#800080"
bordercolordark="#FFFFFF">
        <tr bgcolor="#DFEFFF">
          <td width=90 align=center>������</td>
          <td width=90 align=center>��������</td>
          <td width=270 align=center>����˵��</td>
          <td width=100 align=center>�佱(�ȼ�)</td>
          <td width=60 align=center>��Ʊ��</td>
          <td width=40 align=center>��Ʊ</td><td width=80 align=center>Ͷ��һƱ</td>
<!--#include file="data.asp"-->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
sql="select * from ��԰���ư� order by Ʊ�� desc"
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
          <td width=90 align=center><font color=red><%=rs("������")%></font></td>
          <td width=90 align=center><a href=list.asp?myid=<%=rs("������")%> target=_blank><font color=blue><%=rs("��������")%></font></a></td>
          <td width=270 align=center><%=rs("����˵��")%></td>
          <td width=100 align=center><%if aqjh_grade>=10 then%><a href=jpset.asp?name=<%=rs("������")%>&j=1 target=_blank>[1]</a> <a href=jpset.asp?name=<%=rs("������")%>&j=2 target=_blank>[2]</a> <a href=jpset.asp?name=<%=rs("������")%>&j=3 target=_blank>[3]</a> <a href=jpset.asp?name=<%=rs("������")%>&j=0 target=_blank>[��]</a><%else%>��Ȩ����<%end if%></td>
          <td width=60 align=center>
<%if ytrs=0 then%>
        <div align="center">0%            
          <%else
	n=int(rs("Ʊ��")/ytrs*100)
	%>
          <%=n%>%            
          <%end if%>
</td>
<td width=40 align=center><font color=red><%=rs("Ʊ��")%></font></td><td align=center>
<%if CDate(ks)>now() then%>
ͶƱδ��ʼ
<%elseif CDate(jz)<now() then%>
ͶƱ�ѽ���
<%elseif aqjh_jhdj<dj then%>
�ȼ�����
<%elseif tp=1 then%>
�Ѿ�Ͷ��Ʊ
<%else%>
<a href=tp.asp?name=<%=rs("������")%> target=_blank>Ͷ��һƱ</a>
<%end if%>
</td></tr>
<% 
  rs.movenext
  if rs.eof then exit for
  next
 %>   
��<font color=2020dd><% =rs.pagecount %></font>ҳ<font color=2020dd><% =rs.recordcount %></font>�� ��ǰ:<font color=2020dd><%=page%></font>/<%=rs.pagecount%>&nbsp;&nbsp;<a href=index.asp?pageno=1>��ҳ</a>&nbsp;<a href=index.asp?pageno=<% =page-1 %>>��ҳ</a>&nbsp;<a href=index.asp?pageno=<% =page+1 %>>��ҳ</a>&nbsp;<a href=index.asp?pageno=<% =rs.pagecount %>>ĩҳ</a>
<%end if%>
</table><br>
<center><FONT color=#0000ff>&copy; ��Ȩ���� 2005-2006 </FONT><A href="http://www.7758530.com/" target=_blank><FONT color=#0000ff>���ֽ�����</FONT></A>
</body></html>