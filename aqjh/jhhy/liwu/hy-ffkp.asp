<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
randomize timer
regjm=int(rnd*9998)+1
%>
<html>
<head>
<body bgcolor="#FF9900" background="bg.gif">
<title><%=Application("sjjh_chatroomname")%>��ȡ��������</title>
<link rel="stylesheet" href="../../css.css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#000000" topmargin="2">
<br>
<br>
<center>
  <table width=401 height="227" border=3 align=center cellpadding="5" cellspacing="10" bordercolor="#6633CC" >
      <td width="367" height="252" align="center" valign="middle">���뽭�����ƺ����������ȡ��Ա��Ƭ�� 
        <br> <br>
        <table width="325" height="175" align="center">
          <tr>
            <td height="169"> 
              <form method=POST action='hy-ffkp1.asp' onsubmit='return(check());' name=reg>
                <p align="center">
                  <input type=hidden name=regjm value="<%=regjm%>">
                  <br>
                  ������ 
<input type=text name=name size=12 maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                  <br><br>
                  ���룺 
                  <input type=password name=psw size=12 maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                  <br><br>
                  ��֤�� 
                  <input type=password name=psw2 size=4 maxlength="4" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                  �����룺<font color="#FFFF00"><%=regjm%></font></p>
<p align="center">
                  <input type=submit value=��ȡ��Ƭ name="submit" style="border: 1px solid; font-size: 9pt; border-color:#000000 solid">
				  <input type=button value=�ر� onClick="window.close()" name="button" style="border: 1px solid; font-size: 9pt; border-color:#000000 solid">
</p>
</form>
</td>
</tr>
</table><br>
ÿ����շ���һ��<br>
        ����ֻ����[<font color="#FF0000">����</font>],<font color="#FF0000">��Ա�ȼ�</font>��<br>
        �´�ֻ����[<font color="#FF0000">��Ǯ��</font>],<font color="#FF0000">��Ա�ȼ�</font>��<br>
        �´�ֻ����[<font color="#FF0000">������</font>],<font color="#FF0000">��Ա�ȼ�</font>��<br>
        ʱ��Ϊ�������ڼ���<br>
        2005.03.08<br>
        2005.05.01<br>
        2005.06.01<br>
        2005.07.01<br>
        2005.08.01<br>
        2005.����1<br>
               ����<br>
        [���ֽ���]��ӭ���<a href="../chat/hy.ASP" ><font color="#FFFFFF">�����Ա</font></a>�� 
      </td>
    </table>

</center>
</body>
</html>
<SCRIPT language="JavaScript"> 
<!-- 
var targetwin="jhwindow";
function check()
{
var youname=document.reg.name.value;
var mima1=document.reg.psw.value;
var rz=document.reg.psw2.value
var rz1=document.reg.regjm.value
if(youname=="" || youname==null){window.alert("�㲻������������ô֪���ø�˭ѽ����^_^!");return false;}
if( CheckIfEnglish(youname) )
{
window.alert("�����в����зǷ��ַ���Ӣ�ġ����֣�ֻ��ʹ������Ƭ������^_^!");
return false;
}
if(mima1=="" || mima1==null){window.alert("������������ǲ��У�^_^!");return false;}
if(rz1!=rz){window.alert("��������ȷ����֤�룡");return false;}
return true;
}

function CheckIfEnglish( String )
{ 
    var Letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890`=~!@#$%^&*()_+[]{}\\|/?.>,<;:'\-?<>/�����������������������������������������������������������������ѡ������ ������������ �� �� �������I�ơ� ���� ���ΦƦء� �ӡ��� ���ǅe�� �������� �ɡ衩 ���� �� �ԨK�L ������ੳ ���� ������n���� �� �٢ڢۢܢݢޢߢ� \"";
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