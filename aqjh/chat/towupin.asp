<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Write "<script>window.moveTo(0,0);</script>"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
to1=Trim(Request.QueryString("toname"))
if to1="���" then to1=aqjh_name
if aqjh_grade<9 then
	Response.Write "<script Language=Javascript>alert('������ʲô����');window.close();</script>"
	Response.End
end if
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title>�鿴��Ʒ</title>
<script language="JavaScript">
function s(list){
var lijigongji=liji.checked;
parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();
if (lijigongji==true) {parent.f2.document.af.subsay.click()};
}</script>
<style>
td{font-size:9pt;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed" leftmargin="0" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<p style="font-size:12pt; color:red" align="center"> <font color="#FFFFFF"><b>�� 
  �� �� Ʒ</b><br>
  </font><font color="#FFFFFF"><span style='font-size:9pt'><font size="3" color="#FF0000">��</font><font size="3"><%=to1%><font color="#FF3333">�� 
  </font></font></span></font></p>
<table border="1" cellspacing="0" cellpadding="0" width="325" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#CCCCCC" align="center" >
  <tr> 
    <td width="112" > 
      <div align="center">�� Ʒ ��</div>
    </td>
    <td width="62" > 
      <div align="center">����</div>
    </td>
    <td width="170" > 
      <div align="center">����</div>
    </td>
  </tr>
  <tr bgcolor="#85C2E0" onMouseOut="this.bgColor='#85C2E0';"onMouseOver="this.bgColor='#DFEFFF';"> 
  </tr>
  <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w1,w2,w3,w4,w5,w6,w7,w8,w9,w10 from �û� where ����='"&to1&"'",conn,2,2
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>ҩ Ʒ ��</font></b></div></td>"
if not(isnull(rs("w1"))) then
	call show(rs("w1"),"ʹ��","w1")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>�� ҩ ��</font></b></div></td>"
if not(isnull(rs("w2"))) then
	call show(rs("w2"),"�¶�","w2")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>װ �� ��</font></b></div></td>"
if not(isnull(rs("w3"))) then
	call show(rs("w3"),"Ͷ��","w3")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>�� �� ��</font></b></div></td>"
if not(isnull(rs("w4"))) then
	call show(rs("w4"),"Ͷ��","w4")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>�� Ƭ ��</font></b></div></td>"
if not(isnull(rs("w5"))) then
	call show(rs("w5"),"�ÿ�","w5")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>�� Ʒ ��</font></b></div></td>"
if not(isnull(rs("w6"))) then
	call show(rs("w6"),"��Ʒ","w6")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>�� �� ��</font></b></div></td>"
if not(isnull(rs("w7"))) then
	call show(rs("w7"),"�ͻ�","w7")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>�� ҩ ��</font></b></div></td>"
if not(isnull(rs("w8"))) then
	call show(rs("w8"),"��ҩ","w8")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>ħ �� ��</font></b></div></td>"
if not(isnull(rs("w9"))) then
	call show(rs("w9"),"ħ��","w9")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>�� �� ��</font></b></div></td>"
if not(isnull(rs("w10"))) then
	call show(rs("w10"),"����","w10")
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
sub show(data,lx,lb)
data1=split(data,";")
data2=UBound(data1)
for y=0 to data2-1
	data3=split(data1(y),"|")
%>
  <tr bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="112"> 
      <div align="center"><%=data3(0)%> </div>
    </td>
    <td width="62"> 
      <div align="center"><%=data3(1)%> </div>
    </td>
    <td width="170"> 
	    <div align="center"> <a href="del.asp?cz=1&lb=<%=lb%>&toname=<%=to1%>&name=<%=data3(0)%>" ><font color="#FF0000">ɾ������Ʒ</font></a> 
        <a href="del.asp?cz=2&lb=<%=lb%>&toname=<%=to1%>&name=<%=data3(0)%>"><font color="#0000FF">����10����Ʒ</font></a> </div>
    </td>
  </tr>
  <%
erase data3
next
erase data1
end sub%>
</table>
</body></html>