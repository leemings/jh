<%@ LANGUAGE=VBScript codepage ="936" %>
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
randomize timer
regjm=int(rnd*8998)+1000
id=LCase(trim(request.querystring("id")))
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');window.close();}</script>"
Response.End 
end if
%>
<HTML><HEAD><TITLE>ע���ʺ�</TITLE>
<LINK href="pic/lyy.css" rel=stylesheet>
<SCRIPT language="JavaScript"> 
<!-- 
var targetwin="jhwindow";
function check()
{
var youname=document.reg.name.value;
var mima1=document.reg.psw.value;
var mima2=document.reg.pswc.value;
var e_mail=document.reg.e_mail.value;
var oicq=document.reg.oicq.value;
var regno1=document.reg.regjm1.value;
var regno2=document.reg.regjm2.value;
if(youname=="" || youname==null){window.alert("����Ҫע������֣���^_^!");return false;}
if( CheckIfEnglish(youname) )
{
window.alert("�����в����зǷ��ַ���Ӣ�ġ����֣�ֻ��ʹ������Ƭ������^_^!");
return false;
}
if(mima1=="" || mima1==null){window.alert("������������ǲ��У�^_^!");return false;}
if(mima2=="" || mima2==null){window.alert("��֤���벻�ԣ�^_^!");return false;}
if(mima2!=mima2){window.alert("�����������벻��ͬ��");return false;}
if(regno1!=regno2){window.alert("��������ȷ����֤��:"+regno1+"��");return false;}
return true;
}
function CheckIfEnglish( String )
{ 
    var Letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890`=~!@#$%^&*()_+[]{}\\|/?.>,<;:'\-?<>/�����������������������������������������������������������������ѡ������ ������������ �� �� ������I�ơ� ���� ���ΦƦء� �ӡ��� ���ǅe�� �������� �ɡ衩 ���� �� �ԨK�L ������ੳ ���� ������n���� �� �٢ڢۢܢݢޢߢ� \"";
     var i;
     var c;
     for( i = 0; i < String.length; i ++ )
     {
          c = String.charAt( i );
	  if (Letters.indexOf( c ) < 0)
	     return false;
     }
     return true;
}
//-->
</SCRIPT>
</HEAD>
<BODY background=pic/BG.gif oncontextmenu=self.event.returnValue=false>
<TABLE id=AutoNumber1 style="BORDER-COLLAPSE: collapse" borderColor=#111111 cellSpacing=0 cellPadding=0 width=380 align=center border=0>
<TBODY><TR>
<TD width="3%" height=1><IMG src="pic/T1.gif" border=0></TD>
<TD width="94%" background=pic/M1.gif height=1>��</TD>
<TD width="12%" height=1><IMG src="pic/T2.gif" border=0></TD></TR>
  <TR><TD width="3%" background=pic/M2.gif rowSpan=3></TD>
    <TD width="94%" height=23><P align=center><SPAN lang=en>&copy;   
<B>�������ϸ��д������Ϣ</P></SPAN></b></TD>
    <TD width="12%" background=pic/M2.gif height=1 rowSpan=3></TD></TR>
  <TR>
    <TD width="94%" height=16>
      <HR color=#682420 SIZE=1>
    </TD></TR>
  <TR><TD width="94%">
<!=><form method=POST action='joinjhnow.asp' onsubmit='return(check());' name=reg>
<TABLE height=403 cellSpacing=1 width="99%" border=0>
        <TBODY>
        <TR>
          <TD width="21%" height=25>*������</TD>
          <TD width="31%" height=25><INPUT class=input maxLength=10 size=11 type=text name=name></TD>
          <TD width="48%" height=25>����ʹ��2-5������</TD></TR>
        <TR>
          <TD width="21%" height=25>*�Ա�</TD>
          <TD width="31%" height=25><SELECT size=1 name=sex> 
<OPTION value=�� selected>��</OPTION><OPTION value=Ů>Ů</OPTION></SELECT></TD>
          <TD width="48%" height=25>��ѡ������Ա�</TD></TR>
        <TR>
          <TD width="21%" height=25>*ͷ��</TD>
          <TD width="31%" height=25><select name=face size=1 onChange="document.images['face'].src='../ico/'+options[selectedIndex].value+'-2.gif';">
                    <%for i=1 to 108%>
                    <option value='<%=i%>'><%=i%></option>
                    <%next%>
                  </select></TD>
          <TD width="48%" height=25><IMG id=face src="../ico/n.gif"></TD></TR>
        <TR>
          <TD width="21%" height=25>*���룺</TD>
          <TD width="31%" height=25><INPUT class=input type=password 
            maxLength=10 size=11 name=psw></TD>
          <TD width="48%" height=25>����6-10���ַ�</TD></TR>
        <TR>
          <TD width="21%" height=25>*ȷ�ϣ�</TD>
          <TD width="31%" height=25><INPUT class=input type=password 
            maxLength=10 size=11 name=pswc></TD>
          <TD width="48%" height=25>������һ������ȷ��</TD></TR>
<TR>
          <TD width="21%" height=25>*�ڶ����룺</TD>
          <TD width="31%" height=25><INPUT class=input type=password 
            maxLength=10 size=11 name=psw2></TD>
          <TD width="48%" height=25>�����һض�ʧ������</TD></TR>
<TR>
  
              <td height="25" width="71" align="center">����<font color="#FF0080"> 
              *</font></td>       
              <td height="25" width="276"> &nbsp;        
                    <select name="country"> 
                  <option value="��" selected>������</option>                
                  <option value="�ع�">�ع�</option>
                  <option value="����">����</option>
                  <option value="�Թ�">�Թ�</option>
                  <option value="κ��">κ��</option>
                  <option value="���">���</option>
                  <option value="����">����</option>
                  <option value="���">���</option>
                </select></td>
       </tr>

                 <TD width="21%" height=25>*QQ��</TD>
          <TD width="31%" height=25><INPUT class=input maxLength=11 size=11 name="oicq"></TD>
          <TD width="48%" height=25>��дQQ����������ϵ5λ����</TD></TR>
        <TR>
          <TD width="21%" height=25>*�������䣺</TD>
          <TD width="31%" height=25><INPUT class=input maxLength=25 size=11 
            name=e_mail value="119yejin@163.com"></TD>
          <TD width="48%" height=25>��ȷ��д��ʵ����</TD></TR>        <TR>
          <TD width="21%" height=25>�����ˣ�</TD>
          <TD width="31%" height=25><INPUT class=input maxLength=5 size=11 name=jsr value=""></TD>
          <TD width="48%" height=25>�����ˣ�û��������</TD></TR>
        <TR>
          <TD width="21%" height=25>*��֤��</TD>
          <TD width="31%" height=25><input type=hidden name=regjm1 value="<%=regjm%>"><INPUT class=input maxLength=4 size=11 
            name=regjm2 value='<%=left(regjm,4)%>' onKeypress="if (event.keyCode < 45 || event.keyCode > 57) {event.returnValue = false;}"></TD>
          <TD width="48%" height=25>��֤�����룺<FONT color=red><%=regjm%></FONT></TD></TR>
        <TR>
          <TD align=middle width="100%" colSpan=3 height=29><INPUT style="BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid; BACKGROUND-COLOR: #e8e8d8" type=submit value="����" name="submit"> 
<INPUT style="BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid; BACKGROUND-COLOR: #e8e8d8" onclick=window.close() type=button value="�ر�" name=button></form></TD></TR>
        <TR>
          <TD width="100%" colSpan=3 height=21><FONT 
            color=#ff0000>ע��������μ�������д���ʺź����룬�����Ͻ���</FONT></TD></TR>
        <TR>
          <TD width="100%" colSpan=3 height=17><FONT 
            color=#ff0000>��������������˵����֣��Ա�����ʵ�����</FONT></TD></TR>
        <TR>
          <TD align=middle colSpan=3><FONT 
            color=#0000ff>&copy; ��Ȩ���� 2004-2008 </FONT><a href="http://www.happyjh.com/" target="_blank"><font color="#0000ff">���ֽ�����</font></a> </TD></TR>  
</table><!=></TD></TR>
<TR><TD width="3%" height=6><IMG src="pic/T3.gif" border=0></TD>
    <TD width="94%" background=pic/M3.gif height=6>��</TD>
    <TD width="12%" height=6><IMG src="pic/T4.gif" 
  border=0></TD></TR></TBODY></TABLE></BODY></HTML>