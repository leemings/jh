<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
randomize timer
regjm=int(rnd*9998)+1
%>
<HTML><HEAD><TITLE>�޸�����</TITLE>
<LINK href="pic/lyy.css" rel=stylesheet></HEAD>
<BODY background=pic/BG.gif oncontextmenu=self.event.returnValue=false>
<form method=POST action='modifyok.asp'>
<TABLE id=AutoNumber1 style="BORDER-COLLAPSE: collapse" borderColor=#111111 cellSpacing=0 cellPadding=0 width=380 align=center border=0>
<TBODY><TR>
<TD width="3%" height=1><IMG src="pic/T1.gif" border=0></TD>
<TD width="94%" background=pic/M1.gif height=1>��</TD>
<TD width="12%" height=1><IMG src="pic/T2.gif" border=0></TD></TR>
  <TR><TD width="3%" background=pic/M2.gif rowSpan=3></TD>
    <TD width="94%" height=23><P align=center><SPAN lang=en>&copy;
<B>�� �� �� ��</P></SPAN></TD>
    <TD width="12%" background=pic/M2.gif height=1 rowSpan=3></TD></TR>
  <TR>
    <TD width="94%" height=16>
      <HR color=#682420 SIZE=1>
    </TD></TR>
  <TR><TD width="94%">
<!=>
<table width="254" align="center">
          <tr>
            <td><div align="center">����ID�� 
                  <input type=password name=id size=12 maxlength="10" class=input>
                  <br>
                  �ա����� 
                  <input type=text name=name size=12 maxlength="10" class=input>
                  <br>
                  ԭ���룺 
                  <input type=password name=oldpass size=12 maxlength="10" class=input>
                  <br>
                  �����룺 
                  <input type=password name=pass size=12 maxlength="10" class=input>
                  <br>
                  ȷ���ϣ� 
                  <input type=password name=repass size=12 maxlength="10" class=input>
                </div>
</td>
</tr>
<tr>
<td align="center">
<input type="submit" name="submit" value="�޸�" style="BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid; BACKGROUND-COLOR: #e8e8d8">
<input type="button" value="�ر�" onclick="window.close()"
name="button" style="BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid; BACKGROUND-COLOR: #e8e8d8"></td>
</tr>
<tr>
<td>˵����<br>
����IE5���¼��������룬<br>
������������������ͨ���������ʷ��¼��<br>
��ɾ����¼�������ʺű�����</td></tr>
        <TR>
          <TD align=middle><FONT 
            color=#0000ff>&copy; ��Ȩ���� 2004-2005 </FONT><A 
            href="http://www.7758530.com/" target=_blank><FONT 
            color=#0000ff>���齭����</FONT></A> </TD></TR>
</table>
<!=></TD></TR></form>
<TR>
    <TD width="3%" height=6><IMG src="pic/T3.gif" border=0></TD>
    <TD width="94%" background=pic/M3.gif height=6>��</TD>
    <TD width="12%" height=6><IMG src="pic/T4.gif" 
  border=0></TD></TR></TBODY></TABLE></FORM></BODY></HTML>