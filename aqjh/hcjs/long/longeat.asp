<!--#include file="check_user.asp" -->

<HTML>
<HEAD>
<TITLE>�����̵�</TITLE>
<STYLE type="text/css">
<!--
TD{
  font-size : 12px;
  font-weight : lighter;
  line-height : 16px;
  color : #643706;
  margin-top : 1px;
  margin-left : 1px;
  margin-right : 1px;
  margin-bottom : 1px;
}
A{
  color : #553006;
  text-decoration : none;
}
-->
</STYLE>
</HEAD>
<BODY topmargin="0" leftmargin="0" rightmargin="0" bgcolor="#ddceb9">
<CENTER>
<TABLE border="0" width="550" cellpadding="0" cellspacing="0">
  <TBODY>
    <TR>
      <TD></TD>
    </TR>
    <TR>
      <TD>
      <TABLE border="0" width="100%" cellpadding="0" cellspacing="0">
        <TBODY>
          <TR>
            <TD background="pic/list_image8.gif" width="19" rowspan="2"></TD>
            <TD align="center">
            <TABLE border="0" width="100%" cellpadding="0" cellspacing="0">
              <TBODY>
                <TR>
                  <TD width="23"><IMG src="pic/list_image4.gif" width="23" height="37" border="0"></TD>
                  <TD background="pic/list_image7.gif"><FONT color="#ffffff">�������̵꡿</FONT></TD>
                  <TD width="21"><IMG src="pic/list_image5.gif" width="21" height="37" border="0"></TD>
                  <TD background="pic/list_image6.gif"></TD>
                </TR>
              </TBODY>
            </TABLE>
            <TABLE border="0" width="100%" cellpadding="0" cellspacing="0" height="2">
              <TBODY>
                <TR>
                  <TD bgcolor="#d3c5b1" height="32">&nbsp;<B>���������б�</B></TD>
                </TR>
              </TBODY>
            </TABLE>
            <TABLE width="100%" border="0" cellpadding="0" cellspacing="0">
              <TBODY>
                <TR>
                  <TD bgcolor="#d3c5b1"></TD>
                </TR>
                <TR>
                  <TD bgcolor="#f1efe4">
                    <TABLE bgcolor="#f1efe4" border="0" bordercolor="#000000" cellpadding="1" cellspacing="1" width="100%">
                      <TBODY> 
                      <TR> 
                        <TD bgcolor="#e8e3d0" width="15%" height="20" valign="bottom"> 
                          <DIV align="center"><B>������</B></DIV>
                        </TD>
                        <TD bgcolor="#e8e3d0" width="49%" height="20" valign="bottom"> 
                          <DIV align="center"><B>����</B></DIV>
                        </TD>
                        <TD bgcolor="#e8e3d0" width="18%" height="20" valign="bottom"> 
                          <DIV align="center"><B>�ۼ�</B></DIV>
                        </TD>
                        <TD bgcolor="#e8e3d0" width="18%" height="20" valign="bottom"> 
                          <DIV align="center"><B>����</B></DIV>
                        </TD>
                      </TR>
                      <% 
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="SELECT ��Ʒ��,˵��,����,id FROM ��Ʒ�� where  ����='����'"
rs.open sql,conn,1,1
do while not rs.bof and not rs.eof
%>
                      <TR> 
                        <TD width="15%" bgcolor="#e8e3d0" height="20" valign="bottom"> 
                          <%=rs("��Ʒ��")%></TD>
                        <TD width="49%" bgcolor="#e8e3d0" height="20" valign="bottom"> 
                          <div align="left"><%=rs("˵��")%></div>
                        </TD>
                        <TD width="18%" bgcolor="#e8e3d0" height="20" valign="bottom"> 
                          <DIV align="center">��ң�<%=rs("����")%></DIV>
                        </TD>
                        <TD width="18%" bgcolor="#e8e3d0" height="20" valign="bottom"> 
                          <DIV align="center"><a href="longeat1.asp?id=<%=rs("id")%>">����</a></DIV>
                        </TD>
                        <%
rs.movenext
loop
%>
                      </TR>
                      </TBODY> 
                    </TABLE>
             </TD>
                </TR>
                <TR>
                  <TD align="center" bgcolor="#e8e3d0" height="20" valign="bottom"> 
                    <br>
                    <p><A class="w" href="javascript:window.close()"><B>[�رմ���]</B></A><br>
                    </p>
                  </TD>
                </TR>
              </TBODY>
            </TABLE>
            </TD>
            <TD width="19" background="pic/list_image9.gif" rowspan="2"></TD>
          </TR>
          <TR>
            <TD background="pic/list_mage10.gif"><IMG src="pic/list_mage10.gif" width="12" height="4" border="0"></TD>
          </TR>
        </TBODY>
      </TABLE>
      </TD>
    </TR>
  </TBODY>
</TABLE>
</CENTER>
</BODY>
</HTML>
<%
conn.close
set conn=nothing
set rs=nothing%>
</body>
</html>
