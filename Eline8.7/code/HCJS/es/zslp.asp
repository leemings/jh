<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
if Session("sjjh_inthechat")<>"1" then 
	Response.Write "<script language=JavaScript>{alert('�㲻�ܽ��в��������д˲���������������ң�');window.close();}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
dim page
page=request.querystring("page")
myname=trim(request.querystring("myname"))
PageSize = 15
rs.open "delete * from s where m<now()-5 and c<>'���չ�'",conn,3,3
if myname="" then
rs.open "Select * From s where c='����' Order by m DESC",conn,3,3
else
rs.open "Select * From s where c='����' and  (d='"& myname &"' or b="& sjjh_id&") Order by m DESC",conn,3,3
end if
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>
<head>
<title><%=Application("sjjh_chatroomname")%>-������Ʒ��</title>
<style type="text/css">td           { font-family: ����; font-size: 9pt }
body         { font-family: ����; font-size: 9pt }
select       { font-family: ����; font-size: 9pt }
a            { color: #FFC106; font-family: ����; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: ����; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body leftmargin="0" topmargin="0" bgcolor="#66CCCC">
<table border="0" height="24" width="91%" cellspacing="0" cellpadding="0"
align="center">
<tbody>
<tr>
<td height="15" width="79%" bgcolor="#669999"><font color="#669966"> <font color="#FFFFFF"><b>E��������Ʒ��(</b><font color="#000000">ע��10�콻�ײ��ɹ��Ĺ�����Ʒ���ǽ��͵�����������!</font><b>��<a                    href="zslp.asp?myname=<%=sjjh_name%>">
�����ҵ���Ʒ</a> </b></font><font color="#669966"><font color="#CC0000" face="��Բ"><a href="javascript:this.location.reload()">ˢ�´�ҳ��</a></font></font></font></td>
<td height="15" width="21%" bgcolor="#669999">
<div align="right"><font color="#669966"><font color="#FFFFFF"><b><a href="wupin.asp" target="_blank">�����Լ���Ʒ������</a></b></font></font></div>
</td>
</tr>
</tbody>
</table>
<div align="center">
<table width="91%" align="center" cellspacing="0" border="0"
cellpadding="0">
<tr>
<td width="100%">
<table border="1" cellspacing="1" cellpadding="0" width="720"
bordercolorlight="#EFEFEF">
<tr bgcolor="#FFFFFF">
<td width="7%" height="16">
<div align="center"><font color="#FF6600"> ��Ʒ��</font></div>
</td>
<td width="7%" height="16">
<div align="center"><font color="#FF6600"> ������</font></div>
</td>
<td width="4%" height="16">
<div align="center"><font color="#FF6600"> ����</font></div>
</td>
<td width="9%" height="16">
<div align="center"><font color="#FF6600">����</font></div>
</td>
<td width="17%" height="16">
<div align="center"><font color="#FF6600"> ʱ �� </font></div>
</td>
<td width="31%" height="16">
<div align="center"><font color="#FF6600">˵ ��</font></div>
</td>
<td width="5%" height="16">
<div align="center"><font color="#FF6600"> ԭֵ</font></div>
</td>
<td height="16" colspan="2">
<div align="center"><font color="#FF6600">����</font></div>
</td>
</tr>
<%
count=0
do while not (rs.eof or rs.bof) and count<rs.PageSize
%>
<tr bgcolor="#3399CC" onmouseout="this.bgColor='#3399CC';"onmouseover="this.bgColor='#DFEFFF';">
<td width="7%" height="29">
<div align="center"> <font color="#0000FF"><%=rs("a")%></font>
</div>
</td>
<td width="7%" height="29">
<div align="center"> <%=rs("b")%> </div>
</td>
<td width="4%" height="29">
<div align="center"> <%=rs("j")%> </div>
</td>
<td width="9%" height="29">
<div align="center"><font color="#0000FF"><%=rs("d")%></font></div>
</td>
<td width="17%" height="29">
<div align="center"> <%=rs("m")%> </div>
</td>
<td width="31%" height="29">
<div align="left"></div>
<%=rs("k")%></td>
<td width="5%" height="29">
<div align="center"> <%=rs("n")%> </div>
</td>
<td width="12%" height="29">
<div align="center"><% if sjjh_grade>=9 then%><a href="del.asp?id=<%=rs("id")%>"><font color="#0000FF">ɾ��</font></a><%end if%>
<%if len(rs("k"))>5 then zy=left(rs("k"),4)
if rs("d")=sjjh_name and zy<>"��������" then%>
<a href="zslp1.asp?id=<%=rs("id")%>"><font color="#0000FF"><b>лл��</b></font></a>
<%else%>
<font color="#0000FF"><b>no�ҵ�</b></font>
<%end if%></div>
</td>
<td width="8%" height="29">
<%if len(rs("k"))>5 then zy=left(rs("k"),4)

if rs("d")=sjjh_name and zy<>"��������" then%>
<div align="center"><a href="lpjj.asp?id=<%=rs("id")%>"><font color="#0000FF"><b>��������</b></font></a></div>
<%else%>
<div align="center"><font color="#0000FF"><b>no�ҵ�</b></font></div>
<%end if%>
</td>
</tr>
<%rs.movenext%>
<%count=count+1%>
<%loop
pa=rs.pagecount
mu=musers()
rs.close
conn.close
set rs=nothing
set conn=nothing

%>
</table>
<table border="0" cellspacing="1" cellpadding="1" width="100%" bordercolorlight="#EFEFEF">
<tr>
<td align="left" width="37%" height="2">[��<font color="red"><b><%=pa%></b></font>ҳ][<font
color="red"><b><%=mu%></b></font>��׬��]</td>
<td align="right" width="63%" height="2">
<div align="center">[<a href="zslp.asp?page=<%=page-1%>">��һҳ</a>][��<%=page%>ҳ][<a                    href="zslp.asp?page=<%=page+1%>">��һҳ</a>]</div>
</td>
</tr>
</table>
</table>
</div>

<%
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As ���� from s where c='����'")
musers=tmprs("����")
set tmprs=nothing
if isnull(musers) then musers=0
end function

%>
</body>
