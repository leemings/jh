<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
%>
<html>
<head>
<title><%=Application("aqjh_chatroomname")%>���˼�¼</title>
<style type="text/css">
<!--
p            { line-height: 20px; font-size: 9pt }
table        { font-size: 9pt }
a:link       { color: #FF9900; text-decoration: none }
-->
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body text="#000000" vlink="#0000FF" topmargin="0"
leftmargin="0" background="../bg.gif" alink="#0000FF">
<p align="center"> <font
color="#FF6600"><b><font size="+1">�������˼�¼</font></b></font><font color="#CC0000" face="��Բ"><a href="javascript:this.location.reload()"><br>
  <font color="#0000FF">ˢ��</font></a></font> <br>
  ��л��������Щ�������������ǽ����ģ�<br>
  ֻ���������˴�㹻1000��ſ�����������ʾ������<br>
  <br>
  <font color="#0000FF">��飺���˿����������Լ��ĵ������������������ǽ������˴�����1000ʱ,<br>
  ��ÿ���¶����Դ������ϰ�Ƥһ���������Ǽ�����Զ���5%���㣬����Ӱ�������˵��ݵ㣩�����µ׵�<br>
  ʱ�������ġ������һ���¹�ȥ�ˣ��㽫�ǲ���Ƥ�ģ�ֻ�е�һ�¸����ˣ���ֵ���ۼӣ��������Ҷ����˰ɣ�</font> <br>

<table width="621" border="1" cellpadding="0" cellspacing="0" height="13" align="center" bordercolor="#000000">
  <tr> 
    <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM �û� where allvalue>1000 and ������='"& aqjh_name &"' order by -mvalue",conn
%>
    <td width="28" height="10"> 
      <div align="center"><font color="#000000">ID</font></div>
    </td>
    <td width="67" height="10"> 
      <div align="center"><font color="#000000">����</font></div>
    </td>
    <td width="28" height="10"> 
      <div align="center"><font color="#000000">�Ա�</font></div>
    </td>
    <td width="57" height="10"> 
      <div align="center"><font color="#000000">����</font></div>
    </td>
    <td width="63" height="10"> 
      <div align="center"><font color="#000000">����</font></div>
    </td>
    <td width="125" height="10"> 
      <div align="center">����½</div>
    </td>
    <td width="52" height="10"> 
      <div align="center"><font color="#FF0000">���ݵ�</font></div>
    </td>
    <td width="46" height="10"> 
      <div align="center"><font color="#000000">������</font></div>
    </td>
    <td width="37" height="10"> 
      <div align="center"><font color="#000000">ս����</font></div>
    </td>
    <td width="42" height="10"> 
      <div align="center">��Ƥ��</div>
    </td>
    <td width="52" height="10"> 
      <div align="center"><font color="#000000">��Ƥ</font></div>
    </td>
  </tr>
  <%
jl=0
do while not rs.bof and not rs.eof
jl=jl+1
%>
  <tr> 
    <td width="28" height="30"> 
      <div align="center"><font color="#000000"><%=rs("ID")%></font></div>
    </td>
    <td width="67" height="30"> 
      <div align="center"><font color="#0000FF"><%=rs("����")%></font></div>
    </td>
    <td width="28" height="30"> 
      <div align="center"><font color="#000000"><%=rs("�Ա�")%></font></div>
    </td>
    <td width="57" height="30"> 
      <div align="center"><font color="#000000"><%=rs("����")%></font></div>
    </td>
    <td width="63" height="30"> 
      <div align="center"><font color="#000000"><%=rs("����")%></font></div>
    </td>
    <td width="125" height="30"> 
      <div align="center"><font color="#000000"><%=rs("lasttime")%></font></div>
    </td>
    <td width="52" height="30"> 
      <div align="center"><font color="#FF0000"><%=rs("mvalue")%></font></div>
    </td>
    <td width="46" height="30"> 
      <div align="center"><font color="#000000"><%=rs("grade")%></font></div>
    </td>
    <td width="37" height="30"> 
      <div align="center"><font color="#000000"><%=rs("�ȼ�")%></font></div>
    </td>
    <td width="42" height="30"> 
      <div align="center"><font color="#FF0000"><%=int(rs("mvalue")*0.05)%></font></div>
    </td>
    <td width="52" height="30"> 
      <div align="center"><font color="#000000"> 
        <%
prevtime=CDate(rs("lasttime"))
if DateDiff("m",prevtime,sj)=0 and rs("����2")<>"��Ƥ"&Month(date()) and rs("mvalue")>20 then%>
        <a href="babi.asp?id=<%=rs("id")%>"><font color="#0000FF">��Ƥ</font></a> 
<%else%>
	      ���ܲ��� 
<%end if%>
        </font></div>
    </td>
  </tr>
  <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<p align="center"> <font color="#000000"><br>
��������:<b><font color="#0000FF"><%=(jl)%></font></b><font color="#000000">��</font></font> 
<div align="center"></div>
</body>
</html>