<%@ Language=VBScript %>
<% option explicit%>
<!--#include file="adovbs.asp"-->
<!-- #include file="conn.asp" -->
<%
'###################################
'--- Copyright lanmang (c) S.f.  ---
'---    Http://Www.cidu.net     ---
'main.asp?action=search
'������action=searchΪĬ�ϵ��б���ʾ
'###################################
dim list,strSQL
const MaxPage=10
dim CurrentPage
dim totalPut
dim TotalPages
dim i,j,action
dim copyright
dim Aff,links,link
dim name,chk_name,sex,chk_sex,age,chk_age,mons,days,chk_md,xingzuo,chk_xingzuo,icqnum,chk_icq,likes,chk_likes,doc,chk_doc,address,chk_address,mail,chk_mail,Url,chk_url,mode,chk_mode
Aff = Chr(34)
If Request.Querystring("keys") = "admsearch" then
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�߼�����</title>
<style type="text/css">
<!--
.pt9 {  font: 9pt "����"}
.p9 {  font-family: "����"; font-size: 9pt}
a:link {  text-decoration: none; color: #000000}
a:visited {  text-decoration: none; color: #000000}
a:hover {  text-decoration: underline; color: #FF0000}
table {  font-size: 9pt}
td {  font-size: 9pt}
.unnamed12 {  font-size: 12pt}
#tipBox {position: absolute;
width: 160px;
z-index: 100;
border: 1pt black solid;
font-family:����;
font-size: 9pt;
background: ffffdf;
visibility: hidden}
a:link {color:#000000;text-decoration: none}
a:visited {color:#000000;text-decoration: none}
a:active {color:#000000;text-decoration: none}
a:hover {color:red;text-decoration: none}
-->
</style>

</head>
<body topmargin="0" leftmargin="0">
<form action="admsearch.asp" method="post" align="center">
<input name=action type=hidden value=search>
<input name="work" type="hidden" value="adsearchposted"><input name="user" type="hidden" value="$user">
<div align="center">
<center>
<table bgColor="#ccccff" border="1" borderColorDark="#988cd0" borderColorLight="#988cd0" cellSpacing="0" height="310" width="601">
<tbody>
<tr>
<td colSpan="3" height="15">
<table border="1" borderColorDark="#ffffff" borderColorLight="#000000" cellPadding="4" cellSpacing="0" height="1" width="100%">
<tbody>
<tr bgColor="#f8ebd1">
<td align="middle" height="1" width="100%">
<p align="center">�߼�����
</td>
</tr>
</tbody>
</table>
</td>
</tr>
<tr align="middle" bgColor="#a08cd8">
<td colSpan="3" height="15">
<table border="1" borderColorDark="#ffffff" borderColorLight="#000000" cellPadding="4" cellSpacing="0" height="1" width="100%">
<tbody>
<tr bgColor="#f8ebd1">
<td align="middle" height="1" width="15%"><span class="p9"><a href=main.asp?action=search>��-ͬѧ�б�</a></span></td>
<td height="1" width="15%">
<div align="center">
<p><span class="p9"><a href="add.asp">��-ͬѧ����</a></span></p>
</div>
</td>
<td align="middle" height="1" width="15%">
<div align="center">
<p><span class="p9"><a href="userre.asp">��-�޸�����</a></span></p>
</div>
</td>
<td align="middle" height="1" width="15%">
<div align="center">
<p><span class="p9"><a href="sf.asp?Keys=Login">��-��������</a></span></p>
</div>
</td>
<td align="middle" height="1" width="15%">
<div align="center">
<p><a href="admsearch.asp?keys=admsearch">��-�߼�����</a></p>
</div>
</td>
<td align="middle" height="1" width="25%">
<div align="right">
<p><font class="pt9" color="#996600">����</font><font class="pt9" color="yellow">
</font><font class="pt9" color="#00ffff"><strong><font color="#ff6600">
<!--#include file="zongshu.asp" --> </font></strong></font><font class="pt9" color="yellow">
<font color="#996600">λͬѧ����</font></font></p>
</div>
</td>
</tr>
</tbody>
</table>
</td>
</tr>
<tr align="middle">
<td align="right" height="23" width="137">
<p align="right"><font color="#000000"><font class="pt9">��</font>��������:</font></p>
</td>
<td align="middle" bgColor="#dfe1ff" height="23" width="29"><font color="#000000">
<input name="chk_name" type="checkbox" value="YES">
</font></td>
<td bgColor="#dfe1ff" height="23" width="421">
<div align="left">
<p>
<input maxLength="30" name="name" size="16">
</p>
</div>
</td>
</tr>
<tr align="middle">
<td align="right" height="23" width="137">
<p align="right"><font color="#000000"><font class="pt9">��</font>�༶����:</font></p>
</td>
<td align="middle" bgColor="#dfe1ff" height="23" width="29"><font color="#000000">
<input name="chk_mode" type="checkbox" value="YES">
</font></td>
<td bgColor="#dfe1ff" height="23" width="421">
<div align="left">
<p>
<input maxLength="30" name="mode" size="16">
</p>
</div>
</td>
</tr>
<tr align="middle">
<td align="right" height="24" width="137">
<p align="right"><font color="#000000">���Ա�����:</font></p>
</td>
<td align="middle" bgColor="#dfe1ff" height="24" width="29"><font color="#000000">
<input name="chk_sex" type="checkbox" value="YES">
</font></td>
<td bgColor="#dfe1ff" height="24" width="421">
<div align="left">
<p>
<select class="p9" name="sex" size="1">
<option selected value="">�Ա�ѡ��</option>
<option value="��">��</option>
<option value="Ů">Ů</option>
</select>
</p>
</div>
</td>
</tr>
<tr align="middle">
<td align="right" height="23" width="137">
<p align="right"><font color="#000000">����������:</font></p>
</td>
<td align="middle" bgColor="#dfe1ff" height="23" width="29"><font color="#000000">
<input name="chk_age" type="checkbox" value="YES">
</font></td>
<td bgColor="#dfe1ff" height="23" width="421">
<div align="left">
<p>
<input maxLength="50" name="age" size="4">
��</p>
</div>
</td>
</tr>
<tr align="middle">
<td align="right" height="15" width="137">
<p align="right"><font color="#000000">����������:</font></p>
</td>
<td align="middle" bgColor="#dfe1ff" height="15" width="29"><font color="#000000">
<input name="chk_md" type="checkbox" value="YES">
</font></td>
<td bgColor="#dfe1ff" height="15" width="421">
<div align="left">
<p>
<input maxLength="2" name="mons" size="2">
<font color="#000000">��</font>
<input maxLength="2" name="days" size="2">
<font color="#000000">��</font></p>
</div>
</td>
</tr>
<tr align="middle">
<td align="right" height="23" width="137">
<p align="right"><font color="#000000">��ICQ ����:</font></p>
</td>
<td align="middle" bgColor="#dfe1ff" height="23" width="29"><font color="#000000">
<input name="chk_icq" type="checkbox" value="YES">
</font></td>
<td bgColor="#dfe1ff" height="23" width="421">
<div align="left">
<p>
<input maxLength="10" name="icqnum" size="10">
</p>
</div>
</td>
</tr>
<tr align="middle">
<td align="right" height="23" width="137">
<p align="right"><font color="#000000">���绰����:</font></p>
</td>
<td align="middle" bgColor="#dfe1ff" height="23" width="29"><font color="#000000">
<input name="chk_likes" type="checkbox" value="YES">
</font></td>
<td bgColor="#dfe1ff" height="23" width="421">
<div align="left">
<p>
<input maxLength="50" name="likes" size="30">
</p>
</div>
</td>
</tr>
<tr align="middle">
<td align="right" height="23" width="137">
<p align="right"><font color="#000000">����������:</font></p>
</td>
<td align="middle" bgColor="#dfe1ff" height="23" width="29"><font color="#000000">
<input name="chk_doc" type="checkbox" value="YES">
</font></td>
<td bgColor="#dfe1ff" height="23" width="421">
<div align="left">
<p>
<input maxLength="50" name="doc" size="30">
</p>
</div>
</td>
</tr>
<tr align="middle">
<td align="right" height="23" width="137">
<p align="right">����ϵ��<font color="#000000">ַ����:</font></p>
</td>
<td align="middle" bgColor="#dfe1ff" height="23" width="29"><font color="#000000">
<input name="chk_address" type="checkbox" value="YES">
</font></td>
<td bgColor="#dfe1ff" height="23" width="421">
<div align="left">
<p>
<input maxLength="50" name="address" size="30">
</p>
</div>
</td>
</tr>
<tr align="middle">
<td align="right" height="24" width="137">
<p align="right">�������ʼ�<font color="#000000">����:</font></p>
</td>
<td align="middle" bgColor="#dfe1ff" height="24" width="29"><font color="#000000">
<input name="chk_mail" type="checkbox" value="YES">
</font></td>
<td bgColor="#dfe1ff" height="24" width="421">
<div align="left">
<p>
<input maxLength="50" name="mail" size="30">
</p>
</div>
</td>
</tr>
<tr align="middle">
<td align="right" height="24" width="137">
<p align="right">����ҳ��ַ<font color="#000000">����:</font></p>
</td>
<td align="middle" bgColor="#dfe1ff" height="24" width="29"><font color="#000000">
<input name="chk_url" type="checkbox" value="YES">
</font></td>
<td bgColor="#dfe1ff" height="24" width="421">
<div align="left">
<p>
<input maxLength="50" name="url" size="30">
</p>
</div>
</td>
</tr>
<tr align="middle">
<td align="middle" colSpan="3" height="25">
<input class="p9" name="advsearch" type="submit" value="  �߼�����  ">
&nbsp;&nbsp;&nbsp; <font color="red"><b>�߼�������ѡ�и�ѡ��</b></font></td>
</tr>
</tbody>
</table>
</center>    </div> </form>    </body>    </html>
<%Else%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style>
<!--
a:link {color:#000000;text-decoration: none}
a:visited {color:#000000;text-decoration: none}
a:active {color:#000000;text-decoration: none}
a:hover {color:red;text-decoration: none}
-->
</style>
<title>�� �� �� ��</title>
</head><body topmargin="0" leftmargin="0">


<div align="center"><center><table border="1" width="650" cellspacing="0" cellpadding="3" style="font-family: ����; font-size: 9pt; color: #000000" bordercolor="#988CD0">
<tr><td width="100%" colspan="5"><p align="center"><font size="3" color=red>�� �� �� ��</font></p></td></tr>
<tr>
<td width="20%" bgcolor="#F8EBD1" align="center"><a href=main.asp?action=search>��-ͬѧ�б�</a></td>
<td width="20%" bgcolor="#F8EBD1" align="center"><a href="add.asp">��-ͬѧ����</a></td>
<td width="20%" bgcolor="#F8EBD1" align="center"><a href="userre.asp">��-�޸�����</a></td>
<td width="20%" bgcolor="#F8EBD1" align="center">��-��������</td>
<td width="20%" bgcolor="#F8EBD1" align="center"><a href="admsearch.asp?keys=admsearch">��-�߼�����</a></td>
</tr>
<tr>
<td width="20%" bgcolor="#988CD0">��</td>
<td width="20%" bgcolor="#988CD0">��</td>
<td width="20%" bgcolor="#988CD0">��</td>
<td width="20%" bgcolor="#988CD0">��</td>
<td width="20%" bgcolor="#988CD0">��</td>
</tr>
</table></center></div>
<%'--------------��ʼ׼����������------

if not isempty(request("page")) then
currentPage=cint(request("page"))
else
currentPage=1
end if
'----------------------------
strSQL="SELECT * FROM List WHERE ID LIKE '%%'"
If Request.form("chk_name")="YES" Then ' And Request.form("name")<>"" Then
name   = Request.form("name")
strSQL = strSQL & " and name LIKE '%" & name & "%'"
links  = "&name="&name&"&chk_name=YES"
End if
If Request.form("chk_sex")="YES" Then 'And Request.form("sex")<>"" Then
sex    =Request.Form("sex")
strSQL =strSQL & " and sex LIKE '%" & sex & "%'"
links  = links & "&sex="&sex&"&chk_sex=YES"
End if
If Request.form("chk_age")="YES" Then 'And Request.form("age")<>"" Then
age    =Request.Form("age")
strSQL =strSQL & " and age LIKE '%" & age & "%'"
links  = links & "&age="&age&"&chk_age=YES"
End if
IF request.form("chk_md")="YES" Then 'And Request.form("mons")<>"" Then 'And Request.form("days")<>"" Then
mons    =Request.Form("mons")
days    =request.form("days")
strSQL =strSQL & " and mons LIKE '%" & mons & "%'"
strSQL =strSQL & " and days LIKE '%" & mons & "%'"
links  = links & "&mons="&mons&"&chk_md=YES&days="&days
End if
If Request.form("chk_xingzuo")="YES" Then 'And Request.form("xingzuo")<>"" Then
xingzuo    =Request.Form("xingzuo")
strSQL =strSQL & " and xingzuo LIKE '%" & xingzuo & "%'"
links  = links & "&xingzuo="&xingzuo&"&chk_xingzuo=YES"
End if
If Request.form("chk_icq")="YES" Then 'And Request.form("icqnum")<>"" Then
icqnum    =Request.Form("icqnum")
strSQL =strSQL & " and icqnum LIKE '%" & icqnum & "%'"
links  = links & "&icqnum="&icqnum&"&chk_icq=YES"
End if
if Request.form("chk_likes")="YES" Then 'And Request.form("likes")<>""  Then
likes    =Request.Form("likes")
strSQL =strSQL & " and likes LIKE '%" & likes & "%'"
links  = links & "&likes="&likes&"&chk_likes=YES"
End if
if Request.form("chk_doc")="YES" Then 'And Request.form("doc")<>""  Then
doc    =Request.Form("doc")
strSQL =strSQL & " and doc LIKE '%" & doc & "%'"
links  = links & "&doc="&doc&"&chk_doc=YES"
End if
If Request.form("chk_address")="YES" Then 'And Request.form("address")<>""  Then
address    =Request.Form("address")
strSQL =strSQL & " and address LIKE '%" & address & "%'"
strSQL =strSQL & " and address like '%" & address &"%'"
End if
if Request.form("chk_mail")="YES" Then 'And Request.form("mail") <>"" Then
mail    =Request.Form("mail")
strSQL =strSQL & " and mail LIKE '%" & mail & "%'"
links  = links & "&mail="&mail&"&chk_mail=YES"
End if
if Request.form("chk_url") ="YES" Then 'And Request.form("Url")<>""  THen
url    =Request.Form("url")
strSQL =strSQL & " and url LIKE '%" & url & "%'"
links  = links & "&url="&url&"&chk_url=YES"
End if
if Request.form("chk_mode")="YES" Then 'And Request.form("mode")<>""  THen
mode    =Request.Form("mode")
strSQL =strSQL & " and mode LIKE '%" & mode & "%'"
links  = links & "&mode="&mode&"&chk_mode=YES"
End if
strSQL =strSQL & " Order by ID DESC"

'Response.Write strSQL
Response.Write "<center><font size=2>����������ֵ�ǣ�<font color=red>" &name&sex&age&mons&days&xingzuo&icqnum&likes&doc&address&mail&url&mode&"</font></font></center>"

set list=server.createobject("adodb.recordset")

list.open strSQL,conn,adOpenKeySet,adLockPessimistic
if list.eof and list.bof then
%>
<div align="center"><center>
<table border="1" width="650" cellspacing="0" cellpadding="3" bordercolor="#988CD0" bgcolor="#E0E0E0" style="font-family: ����; font-size: 9pt; color: #000000">
<tr><td width="100%">
<p align="center"><font color="#FF0000">
Sorry! û�������ҵ����ϣ�</font></p>
</td></tr></table></center></div>
<%
searchForm
else
totalPut=list.recordcount
if currentPage=1 then
showContent
showpages
else
if (currentPage-1)*MaxPage<totalPut then
list.move  (currentPage-1)*MaxPage
dim bookmark
bookmark=list.bookmark
showContent
showpages
else
currentPage=1
showContent
searchForm
end if
end if
list.close
end if

set list=nothing

sub searchForm()%>
<form method=post action=main.asp name=searchform >
<div align="center"><center><table border="1" width="650" cellspacing="0" cellpadding="3" bordercolor="#988CD0" bgcolor="#E0E0E0" style="font-family: ����; font-size: 9pt; color: #000000">
<tr><td width="100%"><p align="center">
<select name="kind" style="font-family: ����; font-size: 9pt" size="1">
<option selected value="name">����������</option>
<option value="sex"    >���Ա�����</option>
<option value="age"    >����������</option>
<option value="xingzuo">����������</option>
<option value="icqnum" >��ICQ ����</option>
<option value="likes"  >���绰����</option>
<option value="liuyan" >����������</option>
<option value="address">����ס��ַ</option>
<option value="mail"   >�������ʼ�</option>
<option value="url"    >����ҳ��ַ</option>
<option value="mode"   >���༶����</option>
</select>
<input name="keyword" size="20"  style="font-family: ����; font-size: 9pt">
<input name=action type=hidden value=search>
<input name="ok" style="font-family: ����; font-size: 9pt;BACKGROUND-COLOR: #e7e7de; PADDING-TOP: 1pt" type="submit"
value="<%
if request("action")="search" then
Response.Write "������"
else
Response.Write "�� ��"
end if
%>">
</p></td></tr>
</table></center></div></form>
<%
end sub
sub showcontent
dim cal

cal=1
%>
<div align="center"><center><table border="1" width="650" bgcolor="#E0E0E0" cellspacing="0" cellpadding="3" style="font-family: ����; font-size: 9pt; color: #000099" bordercolor="#988CD0">
<tr>
<td width="71"  align="center">����</td>
<td width="31"  align="center">�Ա�</td>
<td width="37"  align="center">����</td>
<td width="124" align="center">ICQ/OICQ</td>
<td width="91"  align="center">�Ǽ�IP</td>
<td width="183" align="center">�Ǽ�ʱ�� </td>
<td width="57"  align="center">�������</td>
</tr>
<%do while not (list.eof or err)%>
<tr>
<td width="71"  align="center"><a href=lookuser.asp?ID=<%=list("ID")%>><%=list("name")%></a></td>
<td width="31"  align="center"><%=list("sex")    %></td>
<td width="37"  align="center"><%=list("age")    %></td>
<td width="124" align="center"><%=list("mode")   %></td>
<td width="91"  align="center"><%=list("ip")     %></td>
<td width="183" align="center"><%=list("times")  %></td>
<td width="57"  align="center"><%=list("counter")%></td>
</tr>
<% i=i+1
if i>=MaxPage then exit do
list.movenext
loop
%>
</table></center></div>
<%end sub

sub showpages()
dim n
if (totalPut mod MaxPage)=0 then
n= totalPut \ MaxPage
else
n= totalPut \ MaxPage + 1
end if
if n=1 then exit sub
%>
<div align="center"><center><table border="1" width="650" cellspacing="0" cellpadding="3" bordercolor="#988CD0" style="font-family: ����; font-size: 9pt; color: #000000">
<tr><td width="100%">����<%=n%>ҳ��
</center>
<%	   dim k
for k=1 to n
if k=currentPage then
response.write "<b>" & Cstr(k)+" </b>"
else
'                   response.write "<a href='admsearch.asp?"&Session("links")&"&page="+cstr(k)+"'>"+Cstr(k)+"</a> "
response.write "<a href='admsearch.asp?link="&links&"&page="+cstr(k)+"'>"+Cstr(k)+"</a> "
end if
next%>
</td></tr>
<tr>
<td width="100%">
<p align="right">��ǰΪ��<%=currentPage%>ҳ&nbsp; ��һҳ ��һҳ</p>
</td>
</tr>
<%end sub%>
</body></html>
<%END IF%>
