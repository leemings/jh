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
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>
<title><%=Application("sjjh_chatroomname")%>�û�ע��</title>
<link rel="stylesheet" href="../css.css">
</head>
<body bgcolor="#000000" leftmargin="0" topmargin="0">
<center>
<script language="JavaScript">
if (navigator.appName.indexOf("Internet Explorer") != -1)	
document.onmousedown = noright;
function noright()
{
if (event.button == 2 | event.button == 3)
{
	alert("��Ҫ������ ^_^");
	 location.replace("#")
}
}
</script>
  <table border=1 bgcolor="#000000" align=center width=400 cellpadding="3" cellspacing="10" height="500">
    <tr bgcolor="#FFFFFF" align="center">
      <td height="480" bgcolor="#000000"><font color="#FF6600">�������ϸ��д������Ϣ</font>
<table width="325" height="229">
<tr>
            <td height="275"> 
              <form method=POST action='joinjhnow.asp' onsubmit='return(check());' name=reg>
                <p align="left"><font color="#FFFFFF">������</font> 
                  <input type=text name=name size=12 maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #ffffff" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                  <font color="#FF0000">*</font> <font color="#FFFFFF">�û���5������</font> 
                  <input type=hidden name=regjm value="<%=regjm%>">
                  <br>
                  <font color="#FFFFFF">���룺</font> 
                  <input type=password name=psw size=12 maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #ffffff" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                  <font color="#FF0000">*</font><font color="#FFFFFF">����6-10���ַ�</font><br>
                  <font color="#FFFFFF">ȷ�ϣ�</font> 
                  <input type=password name=pswc size=12 maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #ffffff" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                  <font color="#FF0000">*</font><font color="#FFFFFF">ͬ�� </font><br>
                  <font color="#FFFFFF">�Ա�</font> 
                  <select name=sex size=1 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #ffffff" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                    <option value=�� style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #ffffff" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'"selected>�� 
                    <option value=Ů style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #ffffff" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">Ů 
                  </select>
                  <br>
                  <font color="#FFFFFF"> Q Q ��</font> 
                  <input type="text" name="oicq" size="11" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #ffffff" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" maxlength="11">
                  <font color="#FF0000">*</font><font color="#FFFFFF">��дQQ����������ϵ5λ���ϣ�</font> 
                  <br>
                  <font color="#FFFFFF">���䣺</font> 
                  <input type=text name=e_mail size=25 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #ffffff" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="@" maxlength="30">
                  <font color="#FF0000">*</font><font color="#FFFFFF">��ϵʹ��</font><br>
                  <font color="#FFFFFF">�����ˣ�</font> 
                  <input size=10 name=jsr value="<%=id%>"maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #ffffff" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                  <br>
                  <font color="#FFFFFF">[�������������ֻΪϵͳ��¼��û�п��Բ���д]</font><br>
                  <font color="#FFFFFF">����ͷ��</font><img id=face src="../ico/n.gif" alt=�����������> 
                  <select name=face size=1 onChange="document.images['face'].src='../ico/'+options[selectedIndex].value+'-2.gif';" style="BACKGROUND-COLOR: #99CCFF; BORDER-BOTTOM: 1px double; BORDER-LEFT: 1px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000">
                    <option selected value="00">��ѡ��</option>
                    <%for i=1 to 468%>
                    <option value='<%=i%>'><%=i%></option>
                    <%next%>
                  </select> <font color="#FF0000">*</font>
                  <font color="#FFFFFF">ͷ��������������ʾ��</font></p>
				  <input type=hidden name=regjm1 value="<%=regjm%>">
                <font color="#FF0000">��֤�룺</font>��<input type=text name=regjm2 size=4 maxlength="4" onKeypress="if (event.keyCode < 45 || event.keyCode > 57) {event.returnValue = false;}"  style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #ffffff" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                <font color="#FFFFFF">������������ұߵ�����</font><font color="#009933"> 
                <%=regjm%></font> 
                <p align="center">
<input type=submit value=���� name="submit" style="border: 1px solid; font-size: 9pt; border-color:
#000000 solid">
<input type=button value=�ر� onClick="window.close()" name="button" style="border: 1px solid; font-size: 9pt; border-color:
#000000 solid">
</p>
</form>
</td>
</tr>
</table>
        <div align="center"><img src="1.gif" alt="׼���Ѿ����δ���E��ȥ" width="100" height="100"> 
        </div>
        <div align="center">
          <table width="325" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><font color="#CC0000">��ע��</font><font color="#FFFFFF">��</font><font color="#CCCCCC">����̤�㱾������.Ҳ��̤������Ľ�������.������ֻ�ܿ��Լ�.û���˻����.�����Ҫվ�����������.��������ñ�ע��.��Ϊ����û���ܵ��£�</font></td>
            </tr>
          </table>
        </div></table>

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