<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<html>
<head>
<title>ϴ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE>
A{ TEXT-DECORATION: none;color:#ffffff}
.td1 {background-color:#FFFFFF; }
TD {FONT-SIZE: 9pt;cursor:hand; }
.tb2 { font-family: "����", "serif"; font-size:13; line-height:100%;}
</STYLE>
</head>
<body bgcolor="#855CCF" background="../bg.gif" text="#FFFFFF">
<%
namef=Trim(Request.QueryString("name"))
nickname=Session("aqjh_name")
if namef<>"" and namef<> nickname then
  response.write "<script Language='Javascript'>alert('��Ҫ�����˵���!')</script>"
  response.end
end if
grade=Session("aqjh_grade")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="dpk.asp"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'�����ķ�����֧��jet4.0����ʹ�����ص����ӷ�������߳�������
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 
'sql="delete from dbk where datediff('n',logtime,now())>=60 "
'Set Rs=conn.Execute(sql)
sql="select * from dpk where ufrom='"& nickname & "'"
Set Rs=conn.Execute(sql)
if rs.eof and rs.bof then
rs.close
conn.close
response.write "<script Language='Javascript'>alert('������û�в����ƾ֣�������Ķ���ʹ��\""/���˿� 1000 \""����Ҫ�󿪾֡�');location.href='javascript:history.go(-1)';</script>"
response.end
else
dpk=rs("pk")
dpkto=rs("uto")
rs.close
conn.close
end if
dpk2=split(dpk,"|")
dpknumb=ubound(dpk2)
%>
<table border="0" align="center" cellpadding="3" cellspacing="5">
<%
nw=1
dim i
dim writeArr(15)

for i= 0 to dpknumb-1 
nw=paixupk(dpk2(i))
writeArr(nw)= writeArr(nw) & "|" & showpk(dpk2(i))
next 

writePPP=""
for j=1 to 15
writePPP=writePPP & writeArr(j)
next

dpk3=split(mid(writePPP,2),"|")

writestr=""
for i= 0 to dpknumb-1 step 3
writestr=writestr & "<tr>"
writestr=writestr & dpk3(i)
if i+1<dpknumb then writestr=writestr & dpk3(i+1)
if i+2<dpknumb then writestr=writestr & dpk3(i+2)
writestr=writestr & "</tr>" 
next
response.write writestr

%>
</table><BR>
<table borderColorDark='#E1E1E1' class="tb2" border="1" align="center" cellspacing="2" bordercolor="#993300" height="30" cellpadding="4">
  <tr> 
    <td onClick=javascript:s('') onMouseUp="mup(this)" onMouseDown="mdown(this,1)" onMouseOver="mover(this)" onMouseOut="mout(this)"><font color="#FFFFFF">�� ѡ</font></td>
  </tr>
<tr> 
<td onClick=javascript:s('��Ҫ��') onMouseUp="mup(this)" onMouseDown="mdown(this,1)" onMouseOver="mover(this)" onMouseOut="mout(this)"><font color="#FFFFFF">�� ��</font></td>
</tr>
<tr>
<td onClick=javascript:s('������') onMouseUp="mup(this)" onMouseDown="mdown(this,1)" onMouseOver="mover(this)" onMouseOut="mout(this)"><font color="#FFFFFF">��  ��</font></td>
</tr>
<tr> 
<td onclick=javascript:window.open("dpkhelp.htm") height="17" onMouseUp="mup(this)" onMouseDown="mdown(this,1)" onMouseOver="mover(this)" onMouseOut="mout(this)"><font color="#FFFFFF">�� ��</font></td>
</tr>
</table>
<br>
<div align="center"><font size="2">��</font><font size="2"><%=dpknumb%>���ơ�</font><font size="2"><br>
  �������˿ˡ�</font></div><br>
<script language="JavaScript">
parent.m.location.href="about:blank"

function s(list){
var sa=parent.f2.document.af.sytemp.value;
if(list=="������"){parent.sw('[���]');}else{parent.sw('[<%=dpkto%>]');}
if(sa==""||sa.charAt(0)!="/"||sa.charAt(1)!="��"||sa.charAt(2)!="��"||sa.charAt(5)=="��"||list.charAt(0)=="��"||sa.charAt(4)!=" "||sa.charAt(5)=="��"||list.charAt(0)=="��"||list==""||sa=="/����$ "){sa="/����$ "}
else{sa=sa+"\+"}
parent.f2.document.af.sytemp.value=sa+list;
parent.f2.document.af.sytemp.focus();
//window.location.reload();
}
function mover(tb){
tb.borderColorLight='#993300';tb.borderColorDark='#E1E1E1';}
function mdown(tb,tttb){
TimerID=0;
tb.borderColorLight='#f6f6f6';tb.borderColorDark='#000000';
}
function mup(tb){
tb.borderColorLight='#993300';tb.borderColorDark='#E1E1E1';}
function mout(tb){
tb.borderColorLight='#993300';tb.borderColorDark='#E1E1E1';}
var loadtime=500;
</script>
<%
function showpk(pk)
dim wpk
wpk=replace(pk,"��","<img src=dpk/1.GIF><br><font color=#FF0000>")
wpk=replace(wpk,"��","<img src=dpk/2.GIF><br><font color=#000000>")
wpk=replace(wpk,"÷","<img src=dpk/3.GIF><br><font color=#000000>")
wpk=replace(wpk,"��","<img src=dpk/4.GIF><br><font color=#FF0000>")
wpk=replace(wpk,"С��","<img src=dpk/xw.gif><br><br>")
wpk=replace(wpk,"����","<img src=dpk/dw.gif><br><br>")
showpk="<td class=td1 onClick=javascript:s('" & pk & "') align=center> " & wpk & "</font></td>"
end function

function paixupk(pk)
dim wpk
wpk=replace(pk,"��","")
wpk=replace(wpk,"��","")
wpk=replace(wpk,"÷","")
wpk=replace(wpk,"��","")
wpk=replace(wpk,"A","1")
wpk=replace(wpk,"J","11")
wpk=replace(wpk,"Q","12")
wpk=replace(wpk,"K","13")
wpk=replace(wpk,"С��","14")
wpk=replace(wpk,"����","15")
paixupk=wpk
end function
%>
</body></html>