<%@ LANGUAGE=VBScript codepage ="936" %>
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")

randomize timer
regjm=int(rnd*8998)+1000
id=LCase(trim(request.querystring("id")))
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');window.close();}</script>"
Response.End 
end if

if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select �ٻ���1,ת�� from �û� where ����='" & aqjh_name &"'",conn,2,2
zaohuan1=rs("�ٻ���1")
zhuan=rs("ת��")
if zaohuan1<>"��" then
Response.Write "<script language=JavaScript>{alert('����������');window.close();}</script>"
    response.end

	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if
if zhuan<8 then
Response.Write "<script language=JavaScript>{alert('������Ҫ�˴�ת�����ϣ��������ʸ�ô?');window.close();}</script>"
    response.end

	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if

set rs=nothing	
conn.close
set conn=nothing
%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<title><%=Application("aqjh_chatroomname")%>����ע��</title>
<link rel="stylesheet" href="../css.css">
</head>
<body bgcolor="#800000" background="../chat/bg.gif">
<center>
<table border=1 bgcolor="#FFFFFF" align=center width=478 cellpadding="5" cellspacing="10" background="../images/b2.gif" height="438">
<tr bgcolor="#FFFFFF" align="center">
      <td height="525" bgcolor="0088FF" width="444"> <font color="#FF6600">�������ϸ��д������Ϣ</font> 
        <br>
<br>
<table width="366" height="229">
<tr>
            <td height="275" width="358"> 
              <form method=POST action='shenshouok.asp' onsubmit='return(check());' name=reg>
                <p align="left">����������                        
                  <input type=text name=name size=12 maxlength="5" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                  <font color="#FF0000">*</font> ������5������                        
                  <input type=hidden name=regjm value="<%=regjm%>">
                  <br>
                  ��¼���룺                        
                  <input type=password name=psw size=12 maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="wjgy26z">
                  <font color="#FF0000">*</font>Ҫ��6-10���ַ����ٺ�<br>                       
                  ȷ�����룺                        
                  <input type=password name=pswc size=12 maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="wjgy26z">
                  <font color="#FF0000">*</font>ͬ�� <br>                       
                  <br>
                  �ڶ����룺                        
                  <input type=password name=psw2 size=12 maxlength="15" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="aqjhwmhwl">
                  <font color="#FF0000">*</font>Ҫ��10-15���ַ�<br>                       
                  ȷ�����룺                        
                  <input type=password name=pswc2 size=12 maxlength="15" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="aqjhwmhwl">
                  <font color="#FF0000">*</font>ͬ�� <br>�ǳ���Ҫ������ID����ʱ�������õ�¼����<br>                       
                  �Ա�                        
                  <select name=sex size=1 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                    <option value=�� style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'"selected>�� 
                    <option value=Ů style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">Ů 
                  </select>
                  <br>
                  OICQ��                        
                  <input type="text" name="oicq" size="11" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" maxlength="11" value="0000000">
                  &nbsp; <font color="#FF0000">*</font>��дQQ����������ϵ5λ���ϣ� <br>                       
                  ���䣺                        
                  <input type=text name=e_mail size=25 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="cys0831@163.com" maxlength="30">
                  <font color="#FF0000">*</font>��ϵʹ��<br>                       
                  �����ˣ�                                      
                  <input size=10 name=jsr value="���׵���"maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">                  <br>
 <br>
                  ����ͷ��<img id=face src="../ico/n.gif" alt=����������� width="14" height="16"> 
<select name=face size=1 onChange="document.images['face'].src='../ico/'+options[selectedIndex].value+'-2.gif';">         
                    <%for i=1 to 108%>
                    <option value='<%=i%>'><%=i%></option>
                    <%next%>
                  </select>                  ͷ��������������ʾ��</p>                       
                  <input type=hidden name=regjm1 value="<%=regjm%>">��֤�룺<input type=text name=regjm2 size=4 maxlength="4" onKeypress="if (event.keyCode < 45 || event.keyCode > 57) {event.returnValue = false;}"  style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="<%=regjm%>"> ������������ұߵ�����<font color="#ff0000"> <%=regjm%> </font>                       
				
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
var mima2=document.reg.pswc.value;
var mima3=document.reg.psw2.value;
var mima4=document.reg.pswc2.value;
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
if(mima3=="" || mima3==null){window.alert("�ڶ��������Ҫ��һ��Ҫ�У���ID����ʱ���������������õ�¼���룡^_^!");return false;}
if(mima4=="" || mima4==null){window.alert("��֤�ڶ����벻�ԣ�^_^!");return false;}
if(mima3!=mima4){window.alert("�����������벻��ͬ��");return false;}
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
