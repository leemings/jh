<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<html>
<head>
<title>��<%=Application("aqjh_chatroomname")%>ҩƷҩר���꡽</title>
<link rel="stylesheet" href="../../css.css">
</head>

<body topmargin="6"
leftmargin="0" bgcolor="#FFFFFF" background="../JHZB/bg.gif">
<div align="center">
<table width="778" border="0" cellspacing="0" cellpadding="4">
<tr>
<td width="101" align="center" valign="top"><img src="../../images/yq.gif" width="101" height="304"></td>
<td colspan="2" valign="top" align="center">
        <p align="center"> <font color="#6633FF">��<%=Application("aqjh_chatroomname")%>ҩƷҩר���꡽</font></p>
        <font size="+1"><a href="YAOPU.asp"><b>ȥ��ҩƷȥ</b></a></font><br>
        <table border="1" cellspacing="0" bordercolor="#000000" align="center" width="527">
          <tr> 
            <td width="76"> 
              <div align="center">�� ҩ</div>
            </td>
            <td width="271"> 
              <div align="center">�� ��</div>
            </td>
            <td width="60"> 
              <div align="center">�� ��</div>
            </td>
            <td width="56"> 
              <div align="center">�� ��</div>
            </td>
            <td width="62"> 
              <div align="center">�� ��</div>
            </td>
          </tr>
          <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM b where  b='��ҩ' order by h",conn
do while not rs.bof and not rs.eof
%>
          <tr> 
            <form method=POST action='buyduyao.asp?id=<%=rs("id")%>'>
              <td width="76"><div align="center"><%=rs("a")%></div></td> 
              <td width="271"> <div align="center">ʧ����<%=rs("d")%>��ʧ����<%=rs("e")%></div></td>
              <td width="60"><div align="center"><%=rs("h")%>��</div></td>
              <td width="56"> 
                <div align="center">
                  <select name="sl">
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
              <td width="62"> 
                <div align="center"> 
                  <input type="SUBMIT" name="Submit" value="����">
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
      </td>
</tr>
</table>
</div>
</body>
</html>