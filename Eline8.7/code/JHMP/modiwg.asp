<%Response.Expires=0
Response.Buffer=true
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
id=clng(Request("id"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sqlstr="select * from y where id="&id
rs.open sqlstr,conn
if rs.EOF or rs.BOF then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��û����ѡ����书��');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="setup.css">
<title>�书�޸ĳ����wWw.51eline.com��</title></head>
<body text="#000000" background="../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center">�书�޸ĳ���</div>
<form method="post" action="modiwgok.asp" name="setmp" onsubmit="return(check());">
  <table border="1" cellspacing="1" align="center" cellpadding="0" bordercolor="#000000" bgcolor="#006699" width="70%">
    <tr> 
      <td width="67"> 
        <div align="center"><font color="#FFFFFF" size="2">ID</font></div>
      </td>
      <td width="197"><font color="#FFFFFF" size="2"> <%=rs("id")%> </font> 
        <input type="hidden" name="id" value="<%=rs("id")%>" size="15" maxlength="20">
        <font color="#FFFFFF" size="2">�ȼ�:<%=rs("e")%></font> </td>
      <td width="81"> 
        <div align="center"><font size="2" color="#FFFFFF">�书��</font></div>
      </td>
      <td width="183"><font color="#FFFFFF" size="2"> 
        <input type="text" name="a"
value="<%=rs("a")%>" size="15" maxlength="20">
        </font></td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center"><font color="#FFFFFF" size="2">ʹ���书</font></div>
      </td>
      <td width="197"><font color="#FFFFFF" size="2"> 
        <input type="text" name="c"
value="<%=rs("c")%>" size="15" maxlength="3">
        </font></td>
      <td width="81" nowrap> 
        <div align="center"><font color="#FFFFFF" size="2">ʹ������</font></div>
      </td>
      <td width="183"><font color="#FFFFFF" size="2"> 
        <input type="text" name="d"
value="<%=rs("d")%>" size="15" maxlength="3">
        </font></td>
    </tr>
    <tr> 
      <td width="67" nowrap> 
        <div align="center"><font size="2" color="#FFFFFF">����˵��</font></div>
      </td>
      <td colspan="3"><font color="#FFFFFF" size="2"> 
        <input type="text" name="f"
value="<%=rs("f")%>" size="50" maxlength="50">
        </font></td>
    </tr>
    <tr> 
      <td colspan="4"> 
        <div align="center"></div>
        <input type="submit" value="ȷ ��" name="submit">
        <input  onClick="javascript:window.document.location.href='setwg.asp'" value="�� ��" type=button name="back">
        <font color="#FFFFFF" size="-1">��ʽͼƬ�� 
        <select name=g size=1 onChange="document.images['face'].src='gif/'+options[selectedIndex].value+'';" style="BACKGROUND-COLOR: #99CCFF; BORDER-BOTTOM: 1px double; BORDER-LEFT: 1px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000">
   <option value="���" selected>���</option>
          <%for t=1 to 24%>
          <option value="wg<%=t%>.gif">wg<%=t%>.gif</option>
			<%next%>
	    </select>
        <img id=face src="gif/wg1.gif" alt=��ʽͼƬ></font> </td>
    </tr>
    <tr> 
      <td colspan="4" height="29"> <font color="#FFFFFF" size="-1">˵��:����˵��Ϊ��ʹ�ô���ʽʱ������,����##Ϊ������,%%�������ˡ�<br>
        ��ʽ����<font color="#FFFF00">�һ���</font> <br>
        ˵��Ϊ��<font color="#FFFF00">##������������ֻ��˫�Ʒ��졭���һ��ȼ�գ����ص���%%��ȥ����</font></font></td>
    </tr>
  </table>
</form>
<%rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>
<SCRIPT language="JavaScript"> 
<!-- 
var targetwin="add";
var badword = new Array("�侫","��","ȥ��","��ʺ","��","��","����","��","����","������","����","��","ɵ","��","��","�","����","����","��","����","����","����","����","����","����","غ��","��Ƥ","��ͷ","��","�P","��","�H","����","��","��","����","����","����","����","����","��Ů","����","ү","��","�Ҷ�","��","���","���","��","����","����","�ڹ�","��","ʺ","ƨ","����","�׳�","��","���","���˹�","fuck","cao","��","����");
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
var a=document.setmp.a.value;
var f=document.setmp.f.value;
if(a=="" || a==null){window.alert("����Ҫ��ʽ����");return false;}
if(f=="" || f==null){window.alert("����Ҫ˵������");return false;}
if( IsBadWord(a) ){
window.alert("��ʽ��������!");
return false;
}
if( IsBadWord(f) ){
window.alert("��ʽ˵��������!");
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