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
action=request.querystring("action")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT ����,���� FROM �û� WHERE ����='"&aqjh_name&"'",conn
mynl=rs("����")
mytl=rs("����")
rs.close
rs.open "SELECT * FROM w WHERE b=" & aqjh_id & " and i>0 and c='ҩƷ'  order by l ",conn
%>
<html>
<head>
<title>�ҵİ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css">
</head>
<body background="back.gif">
<table width="453" border="0" cellspacing="0" cellpadding="0" height="224" align="left">
  <tr> 
    <td width="65" rowspan="3" valign="top">
      <p>
        <input onClick="javascript:window.document.location.href='index.asp'" title=װ��һ�� type=button value="װ��һ��" name="button7">
        <br>
        <input onClick="javascript:window.document.location.href='daojun.asp'" title=װ��ͷ�� type=button value="װ��ͷ��" name="button">
        <br>
        <input onClick="javascript:window.document.location.href='eat.asp'" title=����ҩ�� type=button value="����ҩ��" name="button62">
        <br>����:<%=aqjh_name%>
        ����:<br><font color=red><%=mynl%></font><br>
      ����:<br><font color=red><%=mytl%></font>
      </p>
    </td>
    <td valign="top" rowspan="3" width="388"> 
      <div align="center"><img src="dao.gif" width="40" height="15">����ҩƷһ��<img src="dao1.gif" width="40" height="15"> 
        <font color="#CC0000" face="��Բ"><a href="javascript:this.location.reload()">ˢ��</a></font></div>
      <div align="center"> ע��:ѡ����ã����������е�ҩƷ��<br>
        <br>
        <table border="1" align="center" width="352" cellpadding="0" cellspacing="0" height="23" bordercolor="#000000">
          <tr align="center"> 
            <td nowrap width="65" height="15"> 
              <div align="center"><font color="#000000">��Ʒ��</font></div>
            </td>
            <td nowrap width="30" height="15"> 
              <div align="center"><font color="#000000">���� </font> </div>
            <td nowrap width="36" height="15"> 
              <div align="center"><font color="#000000">������</font></div>
            </td>
            <td nowrap width="36" height="15"> 
              <div align="center"><font color="#000000">������</font></div>
            </td>
            <td colspan="2" nowrap height="15"> 
              <div align="center"><font color="#000000">��ֵ</font></div>
            </td>
            <td width="67" height="15"> 
              <div align="center"><font color="#000000">ʹ������</font></div>
            </td>
            <td width="71" height="15"> 
              <div align="center"><font color="#000000">����</font></div>
            </td>
          </tr>
          <%
do while not rs.eof
%>
          <tr> 
            <form method=POST action='wupin1.asp?action=<%=rs("a")%>&name=<%=aqjh_name%>'>
              <td width="65" height="10">
                <div align="center"><%=rs("a")%> </div>
              </td>
              <td width="30" height="10"> 
                <div align="center"><font color="#FF0000"><%=rs("i")%> </font> 
                </div>
              <td width="36" height="10"> 
                <div align="center"><font color="#0000FF"><%=rs("e")%></font> 
                </div>
              </td>
              <td width="36" height="10"> 
                <div align="center"><font color="#0000FF"><%=rs("f")%> </font> 
                </div>
              </td>
              <td colspan="2" height="10"> 
                <div align="center"><font color="#FF0000"><%=rs("l")%> </font> 
                </div>
              </td>
              <td width="67" height="10"> 
                <div align="center"> 
                  <select name="wpsl">
                    <option value="1" selected>1</option>
                    <option value="2">2</option>
                    <option value="3">3</option>
                    <option value="4">4</option>
                    <option value="10">10</option>
                    <option value="15">15</option>
                    <option value="20">20</option>
                    <option value="25">25</option>
                    <option value="50">50</option>
                  </select>
                </div>
              </td>
              <td width="71" height="10"> 
                <div align="center"> 
                  <input type="SUBMIT" name="Submit2"  value="����">
                </div>
              </td>
            </form>
          </tr>
          <%
rs.movenext
loop
%>
        </table>
        <br>
        <%
rs.close
rs.open "SELECT * FROM x WHERE b=" & aqjh_id & " and c>0  order by e ",conn
%>
        ����ҩƷ�б�<br>
        <table border="1" align="center" width="350" cellpadding="0" cellspacing="0" height="23" bordercolor="#000000">
          <tr align="center"> 
            <td nowrap width="65" height="15"> 
              <div align="center"><font color="#000000">��Ʒ��</font></div>
            </td>
            <td nowrap width="30" height="15"> 
              <div align="center"><font color="#000000">���� </font> </div>
            <td nowrap width="36" height="15"> 
              <div align="center"><font color="#000000">��Ч</font></div>
            </td>
            <td colspan="2" nowrap height="15"> 
              <div align="center"><font color="#000000">��ֵ</font></div>
            </td>
            <td width="66" height="15"> 
              <div align="center"><font color="#000000">ʹ������</font></div>
            </td>
            <td width="70" height="15"> 
              <div align="center"><font color="#000000">����</font></div>
            </td>
          </tr>
          <%
do while not rs.eof
%>
          <tr> 
            <form method=POST action='xleat.asp?action=<%=rs("a")%>&name=<%=name%>'>
              <td width="65" height="17">
                <div align="center"><%=rs("a")%> </div>
              </td>
              <td width="30" height="17"> 
                <div align="center"><font color="#FF0000"><%=rs("c")%> </font> 
                </div>
              <td width="36" height="17"> 
                <div align="center"><font color="#0000FF"><%=rs("d")%> </font> 
                </div>
              </td>
              <td colspan="2" height="17"> 
                <div align="center"><font color="#FF0000"><%=rs("e")%> </font> 
                </div>
              </td>
              <td width="66" height="17"> 
                <div align="center"> 
                  <select name="xlsl">
                    <option value="1" selected>1</option>
                    <option value="2">2</option>
                    <option value="3">3</option>
                    <option value="4">4</option>
                    <option value="5">5</option>
                    <option value="6">6</option>
                    <option value="7">7</option>
                    <option value="8">8</option>
                    <option value="9">9</option>
                  </select>
                </div>
              </td>
              <td width="70" height="17"> 
                <div align="center"> 
                  <input type="SUBMIT" name="Submit22"  value="����">
                </div>
              </td>
            </form>
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
        <br>
        <input onClick="JavaScript:window.close()" title=�رմ��� type=button value="�رմ���" name="button2">
      </div>
    </td>
  </tr>
  <tr> </tr>
  <tr> </tr>
</table>
</body>
</html>
