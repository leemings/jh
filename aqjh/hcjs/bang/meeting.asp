<%@ LANGUAGE=VBScript%>
<SCRIPT LANGUAGE=JavaScript>if(window.name!='aqjh_win'){var i=1;while(i<=50){window.alert('ˢǮ����ϲ�����𣿵㰡��ˢ������');i=i+1;}top.location.href='../../exit.asp'}</script>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<html>
<head>
<title>���ִ��</title>
<style type="text/css">
td { font-family: ����; font-size: 9pt }
body         { font-family: ����; font-size: 9pt }
select       { font-family: ����; font-size: 9pt }
a            { color: #FFC106; font-family: ����; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: ����; font-size: 9pt; text-decoration: 
               underline }
</style>
</head>
<body bgcolor="#000000" text="#FFFFFF">
<table border="0" height="24" width="623" cellspacing="0" cellpadding="0" align="center">
  <tbody> 
  <tr> 
    <td height="38" width="623"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0"
              bordercolorlight="#000000" bordercolordark="#FFFFFF"
              align="center">
        <tr> 
          <td width="91" height="26"><font size="2">&nbsp; <font
                    color="#FFFFFF"></font><font size="2"><font color="#ffffff"
                    size="2"><span class="zilong"><font color="#FFCC00"> <%
y=Month(date())
r=Day(date())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
Response.Write  y & "��" & r & "��" %> </font></span></font></font></font></td>
          <td width="475" height="26"> 
            <div align="center"> </div>
          </td>
          <td width="104"></div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  </tbody> 
</table>
<table width="623" border="0" cellspacing="0" cellpadding="0" align="center" height="298">
  <tr>
    <td width="12" height="305">&nbsp;</td>
    <td width="37" valign="top" height="305"> 
      <div align="center"> </div>
    </td>
    <td valign="top" width="549" height="305"> 
      <div align="center">
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="303">
          <tr> 
            <td align="center" height="47"> <img src="images/dtitle2.gif" width="200" height="45"></td>
            <td height="47"> 
              <div align="right"></div>
            </td>
          </tr>
          <tr valign="top" align="center"> 
            <td colspan="2" height="249"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0" height="141">
                <tr> 
                  <td width="11">&nbsp;</td>
                  <td height="13" class="td01">&nbsp;</td>
                  <td width="11">&nbsp;</td>
                </tr>
                <tr> 
                  <td class="td">&nbsp;
                  </td>
                  <td valign="top"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="0" height="125">
                      <tr> 
                        <td align="center" height="35" colspan="2"><img src="images/gold.gif" width="120" height="40"></td>
                        <td height="35" align="center" colspan="2"><img src="images/silver.gif" width="120" height="40"></td>
                        <td height="35" align="center" colspan="2"><img src="images/copper.gif" width="120" height="40"></td>
                      </tr>
                      <tr> 
                        <td width="16%" height="22" align="center">״Ԫ�� 
                        <%
			set conn=server.createobject("adodb.connection")
                        conn.open Application("aqjh_usermdb")
                        set rs=server.CreateObject ("adodb.recordset")
                        rs.open "select * from ���",conn%>
                        <%=rs("����")%>
                        </td>
                        <td width="15%" height="22" align="center"><%if rs("����")<>aqjh_name then%><a href='fight.asp?typename=gold&id=<%=rs("ID")%>'><%end if%><img src="images/fight.gif" width="80" height="30" border="0"></a> 
                        <%rs.movenext%>
                        </td>
                        <td width="17%" height="22" align="center">״Ԫ�� 
                         <%
                         set rs1=server.createobject("adodb.recordset")
			 rs1.open "select * from ����",conn
			 %>
                          <%=rs1("����")%> 
                        </td>
                        <td width="15%" height="22" align="center"><%if rs1("����")<>aqjh_name then%><a href='fight.asp?typename=silver&id=<%=rs1("ID")%>'><%end if%><img src="images/fight.gif" width="80" height="30" border="0"></a> 
                          <%rs1.movenext%>
                        </td>
                        <td width="16%" height="22" align="center">״Ԫ�� 
                          <%
                          set rs2=server.createobject("adodb.recordset")
			  rs2.open "select * from ͭ��",conn
			  %>
                          <%=rs2("����")%> 
                        </td>
                        <td width="21%" height="22" align="center"><%if rs2("����")<>aqjh_name then%><a href='fight.asp?typename=copper&id=<%=rs2("ID")%>'><%end if%><img src="images/fight.gif" width="80" height="30" border="0"></a> 
                          <%rs2.movenext%>
                        </td>
                      </tr>
                      <tr> 
                        <td width="16%" height="23" align="center">���ۣ�<%=rs("����")%> 
                        </td>
                        <td height="23" align="center" width="15%"><%if rs("����")<>aqjh_name then%><a href='fight.asp?typename=gold&id=<%=rs("ID")%>'><%end if%><img src="images/fight.gif" width="80" height="30" border="0"></a> 
                          <%rs.movenext%>
                        </td>
                        <td width="17%" height="23" align="center">���ۣ�<%=rs1("����")%> 
                        </td>
                        <td width="15%" height="23" align="center"><%if rs1("����")<>aqjh_name then%><a href='fight.asp?typename=silver&id=<%=rs1("ID")%>'><%end if%><img src="images/fight.gif" width="80" height="30" border="0"></a> 
                          <%rs1.movenext%>
                        </td>
                        <td width="16%" height="23" align="center">���ۣ�<%=rs2("����")%> 
                        </td>
                        <td width="21%" height="23" align="center"><%if rs2("����")<>aqjh_name then%><a href='fight.asp?typename=copper&id=<%=rs2("ID")%>'><%end if%><img src="images/fight.gif" width="80" height="30" border="0"></a> 
                          <%rs2.movenext%>
                        </td>
                      </tr>
                      <tr> 
                        <td width="16%" height="22" align="center">̽����<%=rs("����")%></td>
                        <td height="22" align="center" width="15%"><%if rs("����")<>aqjh_name then%><a href='fight.asp?typename=gold&id=<%=rs("ID")%>'><%end if%><img src="images/fight.gif" width="80" height="30" border="0"></a></td>
                        <td width="17%" height="22" align="center">̽����<%=rs1("����")%></td>
                        <td width="15%" height="22" align="center"><%if rs1("����")<>aqjh_name then%><a href='fight.asp?typename=silver&id=<%=rs1("ID")%>'><%end if%><img src="images/fight.gif" width="80" height="30" border="0"></a></td>
                        <td width="16%" height="22" align="center">̽����<%=rs2("����")%></td>
                        <td width="21%" height="22" align="center"><%if rs2("����")<>aqjh_name then%><a href='fight.asp?typename=copper&id=<%=rs2("ID")%>'><%end if%><img src="images/fight.gif" width="80" height="30" border="0"></a></td>
                      </tr>
                    </table>
                  </td>
                  <td class="td">&nbsp;</td>
                </tr>
                <tr> 
                  <td height="13">&nbsp;</td>
                  <td height="13" class="td01">&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
              </table>
              <br>
              <table width="80%" border="0" cellspacing="0" cellpadding="0" height="49">
                <tr>
                  <td height="60"><font color="#FF0000">��������</font>�� ��λ������Ĵ�������ͭ�����书<font color="#FF0000">300</font>���£����������书<font color="#FF0000">300��800</font>���ҽ�����书<font color="#FF0000">800</font>���ϡ��ڽҰ�ǰ����˼�����У��Ұ�ɹ����Խ��������������������֪ͨ��ң�����Ϊ�����˵���ս���󣬻���һ���Ľ������������񲻳ɹ�����һ������ʧ��������ʲô,����һ�Ա�֪��^_^</td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </div>
    </td>
    <td width="25" height="305">&nbsp;</td>
  </tr>
</table>
<div align="center">
</div>
</body>
</html>
<%
rs.close
set rs=nothing
rs1.close
set rs1=nothing
rs2.close
set rs2=nothing
conn.close
set conn=nothing
%>