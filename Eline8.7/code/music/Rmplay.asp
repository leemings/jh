<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if session("sjjh_name")="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ�����������Ƚ��뽭�������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT ����,��� FROM �û� WHERE ����='"& session("sjjh_name") &"'",conn,1,3
if rs("����")<200000 or rs("���")<2 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "����û200000�������Ҳ���2������������RM��ʽ�ĸ�����"
	response.end
end if
rs("����")=rs("����")-200000
rs("���")=rs("���")-2
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
%>                                                                                                                                                                                                                                                                                                                                                                                                                                                                    <script>if(top==self)top.location="1" </script>
<!--#include file="conn.asp"-->
<!--#include file="function.asp"-->
<%
if request("id")<>"" then
set rs=server.createobject("adodb.recordset")
id=request("id")
sql="select * from MusicList where id="&id
rs.open sql,conn,1,3
if rs.eof then
errmsg="<li>�Բ��𣡸ø��������ڣ������Ѿ�������Աɾ����</li>"
call error()
Response.End 
else
rs("hits")=rs("hits")+1
rs.Update
Rm=rs("Rm")
hits=rs("hits")
Singer=rs("Singer")
MusicName=rs("MusicName")
end if
rs.Close

else
errmsg= "<li>��ѡ�������</li>"
call error()
Response.End 
end if

set rs=nothing
conn.close
set conn=nothing
%>

<html>

<head>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<title>һ������--RM��ʽ��������</title>

<style type="text/css">

 <!--

 td {  font-size: 9pt; color: #000000}

 a:link {  font-size: 9pt; color:#000000; text-decoration: none}

 a:active {  font-size: 9pt; ; text-decoration: underline}

 a:visited {  font-size: 9pt; color: #000000; text-decoration: none}

 a:hover {  font-size: 9pt; color:#000000; text-decoration: underline; background-color: #00FFFF}

 -->

 </style>

<meta name="GENERATOR" content="Microsoft FrontPage 3.0">

</head>

<BODY bgColor=#ffffff leftMargin=0  topMargin=0 oncontextmenu="return false" ondragstart="return false" onselectstart="return false" background="image/rm_bg.gif">

<p align="left"><br></p>

<p align="left">&nbsp; </p>

<table width="98%" border="0" cellspacing="1">

<tr> 

<td width="7%">&nbsp;</td>

<td width="33%"> 

<p align="left">��ʽ�� RM</p>

<p align="left">���֣�<%=Singer%></p>

<p align="left">������<%=MusicName%></p>

<p align="left">������<%=hits%>�� </p>

<br>

<p align=left><a href=http://www.k666.com target=NEW title=���˸�����������,���Щ������Ա����,���ǻᾡ����>�������</a></p>

</td>

<td width="60%">

            <OBJECT classid="clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA" height=60 id=video1 style="LEFT: 0px; TOP: 0px" width=238 VIEWASTEXT>
              <param name="_ExtentX" value="5556">
              <param name="_ExtentY" value="1588">
              <param name="AUTOSTART" value="-1">
              <param name="SHUFFLE" value="0">
              <param name="PREFETCH" value="0">
              <param name="NOLABELS" value="0">
              <param name="SRC" value="<%=Rm%>">
              <param name="CONTROLS" value="StatusBar,ControlPanel">
              <param name="CONSOLE" value="RAPLAYER">
              <param name="LOOP" value="0">
              <param name="NUMLOOP" value="0">
              <param name="CENTER" value="0">
              <param name="MAINTAINASPECT" value="0">
              <param name="BACKGROUNDCOLOR" value="#000000"><embed src="<%=Rm%>" type="audio/x-pn-realaudio-plugin" console="Clip1" controls="ControlPanel,StatusBar" height="60" width="238" autostart="true"> 
            </OBJECT>

</td>

</td>

</tr>

<tr> 

<td width="7%">&nbsp;</td>

<td width="33%"> 

<p align="left">&nbsp;</p>

</td>

<td width="60%">&nbsp;</td>

</tr>

</table>

<p>&nbsp;</p>

</BODY></HTML>
