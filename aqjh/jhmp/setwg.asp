<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
id=LCase(trim(request.querystring("id")))
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM �û� where ����='"&aqjh_name&"'",conn
if rs.bof or rs.eof then
	rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Redirect "error.asp?id=453"
	response.end
end if
if rs("����")<>"����" then
	rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Write "<script language=JavaScript>{alert('�������ţ���Ҫ�Ҵ���');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
pai=rs("����")
rs.close
%>
<html>
<head>
<title><%=pai%>�书����</title>
<LINK href="../css.css" rel=stylesheet>
</head>
<body background="../jhimg/bk_hc3w.gif">
<div align="center"><font color="#009900"><b>[ <%=pai%> ] �� �� �� ��</b></font><br>
  <font color="#000000" size="2"><a href="addwg.asp"><b>������ʽ�޸�</b></a></font><br>
</div>
<table cellpadding="1" cellspacing="0" border="1" align="center" width="74%" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr> 
    <td height="25" width="107"> 
      <div align="center"> �� ʽ �� </div>
    </td>
    <td height="25" width="130"> 
      <div align="center">�� �� �� ��</div>
    </td>
    <td height="25" width="141"> 
      <div align="center"> �� �� �� �� </div>
    </td>
    <td height="25" width="97"> 
      <div align="center">�� ��</div>
    </td>
    <td height="25" width="68"> 
      <div align="center"> ���� </div>
    </td>
  </tr>
  <%
rs.open "SELECT * FROM y where b='" & pai & "' order by (c+d)",conn
do while not rs.eof and not rs.bof
%>
    <tr> 
      
    <td height="2" width="107"> 
      <div align="center"> <font color="#000000"><%=rs("a")%></font> <font color="#000000"> 
          <input type="hidden" name="id" size="10" value="<%=rs("id")%>" maxlength="10">
          </font></div>
      </td>
      
    <td height="2" width="130"> 
      <div align="center"> <font color="#000000"><%=rs("c")%></font></div>
      </td>
      
    <td height="2" width="141"> 
      <div align="center"> <font color="#000000"><%=rs("d")%></font></div>
      </td>
      
    <td height="2" width="97"> 
      <div align="center"><font color="#000000"> <%=int(sqr((rs("e")/100)))+1%> 
          </font></div>
      </td>
      <td height="2" width="68"> 
        <div align="center"> <font color="#000000" size="2"><a href="modiwg.asp?id=<%=rs("id")%>">�޸�</a></font> 
          <font color="#000000" size="2"><a href="del.asp?id=<%=rs("id")%>">ɾ��</a></font> 
        </div>
      </td>
    </tr>
    <tr> 
      
    <td height="5" colspan="5"> <font color="#0000FF">˵����<%=rs("f")%><br>
    <%if rs("g")<>"���" then%>
      <img src="gif/<%=rs("g")%>">
      <%else%>���ͼƬ<%end if%> </font></td>
    </tr>
  <%
rs.movenext
loop
rs.close
%>
</table>
<br>
<%
rs.open "SELECT * FROM p where a='" & pai & "'",conn
%>
<form method=POST action='setmp.asp' name=setmp onsubmit='return(check());'>
<table width="67%" border="1" align="center" bgcolor="#3399CC" bordercolor="#000000">
  <tr> 
    <td width="24%" height="30"> 
      <div align="center">�������֣�</div>
    </td>
    <td width="76%" height="30"> 
      <input type="text" name="a"  readonly size="15" maxlength="10" value="<%=rs("a")%>">
    </td>
  </tr>
  <tr> 
    <td width="24%"> 
      <div align="center">����˵����</div>
    </td>
    <td width="76%"> 
      <input type="text" name="e" size="40" value="<%=rs("e")%>" maxlength="45">
      (��������������<font color="#FFFF00"></font>)</td>
  </tr>
  <tr> 
    <td width="24%"> 
      <div align="center">�������ƣ�</div>
    </td>
    <td width="76%"> 
      <input type="text" name="g" size="40" value="<%=rs("g")%>" maxlength="45">
      (��������<font color="#FFFF00">True</font>)</td>
  </tr>
  <tr> 
    <td width="24%"> 
      <div align="center">���ɼ�飺</div>
    </td>
    <td width="76%"> 
        <input type="text" name="d" size="40" value="<%=rs("d")%>" maxlength="45">
      (���ҽ���) </td>
  </tr>
  <tr>
    <td colspan="2" height="15"> 
      <div align="center">
        <input type=submit value=�޸� name="submit" style="border: 1px solid; font-size: 9pt; border-color:
#000000 solid">
      </div>
    </td>
  </tr>
  <tr> 
    <td colspan="2" height="30"> ˵��������˵��Ϊ�������ɵ�����˵�����������ʽ���Ӧ�ġ�<br>
      �磺����ս���ȼ���3�����ϵļ���,����ʽΪ��<font color="#FFFF00">�ȼ�&gt;=3<br>
      </font>�磺�����ݶ��ȼ�����2����Ů����Ҽ��룬����ʽΪ��<font color="#FFFF00">�ȼ�&gt;=2 and �Ա�='Ů'</font><br>
      �磺ֻ���ѻ���ң����´���1000�Ĳſɼ��룬����ʽΪ��<font color="#FFFF00">��ż&lt;&gt;'��' and ����&gt;1000</font><br>
      �磺��������1�����������1��ģ�����ʽΪ��<font color="#FFFF00">����&gt;10000 or ����&gt;10000</font><br>
      <font color="#00FF00">�磺������룬�����룺<b>True</b>(ע��������ĸ����ȫΪ���)</font><br>
      <b>����Ҫ��Ҫ��ȥ���ã�������ò���ȷ�᲻�ܼ������ɡ�</b></td>
  </tr>
</table>
  </form>
<%rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<p class="p1" align="center">ɱ����=<b><font color="#FF0000">(</font>(</b><font color="#0000FF">�Լ��Ĺ���-�Է��ķ���</font><b>)</b>+��ʽ������+��ʽ���书<b><font color="#FF0000">)</font></b>x���Եȼ�<br>
  <br>
  ����书���ֲ����ţ��ٸ�������ɾ����Ӧ���ɣ������ָ���<br>
  <br>
  �书������Ϊɱɱ������Ҫ���أ����Ը���ʵ���������ѡ�����Ϊ��С���������<br>
  <br>�ڿ��ֽ�����վ��</p>
</body>

</html>
<SCRIPT language="JavaScript"> 
<!-- 
var targetwin="add";
var badword = new Array("�侫","��","ȥ��","��ʺ","��","��","����","��","����","������","����","��","��","��","�","����","����","��","����","����","����","����","����","����","غ��","��Ƥ","��ͷ","��","�P","��","�H","����","��","��","����","����","����","����","����","��Ů","����","ү","��","�Ҷ�","��","���","���","��","����","����","�ڹ�","��","ʺ","ƨ","����","�׳�","��","���","���˹�","fuck","cao","��","����");
var badstr = "~!@ #$%^&*()[]{}_+-|=\`;,:'\"?<>/������������������������������������������������������������" ;
function IsBadWord(m)
{var tmp = "" ;
for(var i=0;i<m.length; i++)
{for(var j=0;j<badstr.length;j++)
if(m.charAt(i) == badstr.charAt(j)) break;
if(j==badstr.length) tmp += m.charAt(i) ;}
for(i=0;i<badword.length;i++) if(tmp.search(badword[i]) != -1) return true;
return false;}
function check()
{
var e=document.setmp.e.value;
var g=document.setmp.g.value;
var d=document.setmp.d.value;
if(e=="" || e==null){window.alert("����Ҫ˵������");return false;}
if(g=="" || g==null){window.alert("����Ҫ��������");return false;}
if(d=="" || d==null){window.alert("����Ҫ�������");return false;}
if( IsBadWord(e) ){
window.alert("˵��������!");
return false;
}
if( IsBadWord(g) ){
window.alert("����������!");
return false;
}
if( IsBadWord(d) ){
window.alert("��鲻����!");
return false;
}
}
function check1()
{
var youname=document.modi.wgname.value;
if(youname=="" || youname==null){window.alert("�������书���֣���^_^!");return false;}
if( IsBadWord(youname) ){
window.alert("��ʽ���ֲ�����!");
return false;
}
if( CheckIfEnglish(youname) )
{
window.alert("�书�в����зǷ��ַ���Ӣ�ġ����֣�ֻ��ʹ������Ƭ������^_^!");
return false;
}
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