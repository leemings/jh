<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
%>
<html>
<head>
<title>���ֽ�������</title>
<link rel="stylesheet" href="../../css.css">
</head>
<body topmargin="6" leftmargin="0" bgcolor="#FFFFFF" background="bk2.gif" oncontextmenu=self.event.returnValue=false>
<center>
<br><div align="center"><font size="5" face="����" color=green>�� �� �� ��</font><BR><BR>[<a href=javascript:window.close()>��Щ�����������ô���꣬ʵ���᲻�õ�����</a>]<br><br>
<table width="530">
<tr>
    <td colspan="2" valign="top" align="center">    
      <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="100%">
        <tr> 
          <td width="80" height="13"> 
            <div align="center">��Ʒ��</div>
          </td>
          <td width="80" height="13"> 
            <div align="center">����</div>
          </td>
          <td width="60" height="13"> 
            <div align="center">����</div>
          </td>
          <td width="100" height="13"> 
            <div align="center">ԭ��</div>
          </td>
          <td width="100" height="13"> 
            <div align="center">�ּ�</div>
          </td>
          <td width="55" height="13"> 
            <div align="center">����</div>
          </td>
          <td width="51" height="13"> 
            <div align="center">����</div>
          </td>
        </tr>
<%
rs.open "SELECT w1,w2,w4,w5,w6,w7 FROM �û� WHERE ����='" & aqjh_name& "'",conn
if IsNull(rs("w1")) and IsNull(rs("w2")) and IsNull(rs("w4")) and IsNull(rs("w5")) and IsNull(rs("w6")) and  IsNull(rs("w7")) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "<tr><td colspan=7>�����ϰ循�е���������һ�����У��������뵱�˰ɣ�</td></tr></table>"
	response.end
end if
d5=rs("w5")
d1=rs("w1")
d2=rs("w2")
d4=rs("w4")
d6=rs("w6")
d7=rs("w7")
rs.close
set rs=nothing
conn.close
set conn=nothing
if not IsNull(d5) then
	call show(d5,"��Ƭ")
end if
if not IsNull(d1) then
	call show(d1,"ҩƷ")
end if
if not IsNull(d2) then
	call show(d2,"��ҩ")
end if
if not IsNull(d4) then
	call show(d4,"����")
end if
if not IsNull(d6) then
	call show(d6,"��Ʒ")
end if
if not IsNull(d7) then
	call show(d7,"�ʻ�")
end if
sub show(data,lx)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
data1=split(data,";")
data2=UBound(data1)
for y=0 to data2-1
	data3=split(data1(y),"|")
	rs.open "SELECT h FROM b WHERE a='" & data3(0)& "'",conn
	if not rs.bof and not rs.eof then
%>
        <tr> 
          <form method=POST action='dan1.asp?wpname=<%=data3(0)%>&lx=<%=lx%>'>
	<input type=hidden name=allsl value='<%=data3(1)%>'>
            <td> 
              <div align="center"><%=data3(0)%></div></td>
            <td> 
              <div align="center"><%=lx%></div>
            </td>
            <td> 
              <div align="left"><%=data3(1)%></div>
            </td>
            <td> 
              <div align="left"><%=rs("h")%><%if lx="��Ƭ" then%>Ԫ<%else%>��<%end if%>/��</div>
            </td>
            <td> 
              <div align="center"><%if lx="��Ƭ" then%><%=int(rs("h")/4)%>Ԫ<%else%><%=int(rs("h")/3)%>��<%end if%>/��</div></td>
            <td> 
              <div align="center"><input name="dansl" size=4 value="<%=data3(1)%>">
              </div>
            </td>
            <td> 
              <div align="center"> 
                <input type="submit" name="submit" value="����">
              </div>
            </td>
          </form>
        </tr>
  <%
  end if
  rs.close
erase data3
next
erase data1
set rs=nothing
conn.close
set conn=nothing
end sub
%>
</table>
<p align=center><FONT color=#0000ff>&copy; ��Ȩ���� 2015-2015 </FONT><A href="http://cys.hxjhw.com/" target=_blank><FONT color=#0000ff>���ֽ�����</FONT></A></p>
</body>