<%Response.Expires=0
Response.Buffer=true
un=session("yx8_mhjh_username")
if un="" then Response.Redirect "../../error.asp?id=016"
id=clng(Request("id"))
%><!--#include file="data.asp"--><%
rs.open "select ����,id,ͼ from mj where id="& id &" and ������='"&un&"' and ����='��'",connb
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	connb.close
	set connb=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����û�������书��');history.go(-1);</script>"
	response.end
else
mm=rs("����")
id=rs("id")
tu=rs("ͼ")
end if
%><head>
<link rel=stylesheet href='../../style.css'>
<title>�޸��޼���</title></head>
<body background='../../chatroom/bg1.gif'>
<div align="center">�޸��޼���</div>
<form method="post" action="mod2.asp" name="setmp" onsubmit="return(check());">
  <table border="1" cellspacing="1" align="center" cellpadding="0" bordercolor="#000000" width="397">
    <tr> 
      <td width="65"> 
        <div align="center">ħ��ս��</div>
      </td>
      <td width="134"><%=un%></td>
      <td width="67"> 
        <div align="center">�书��</div>
      </td>
      <td width="111"> 
        <input type="text" name="a"
value="<%=mm%>" size="14" maxlength="4">
         <input type="hidden" name="id" value="<%=id%>" size="15" maxlength="20">
        </td>
    </tr>
    <tr> 
      <td width="389" colspan="4">  
        <select name=c size=1 onChange="document.images['face'].src='gif/'+options[selectedIndex].value+'';" style="BACKGROUND-COLOR: #99CCFF; BORDER-BOTTOM: 1px double; BORDER-LEFT: 1px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000">    
          <option value="���" selected>��ʽͼƬ</option>
          <option value="���">���ͼƬ</option>
         <%for t=1 to 24%>
          <option value="wg<%=t%>.gif">wg<%=t%>.gif</option>
			<%next%>
		 </select>     
        <img id=face src="gif/<%=tu%>" alt=���ͼƬ>   
      </td>  
    </tr>  
    <tr>   
      <td colspan="4" width="389">     
       <p align="center">    
       <input type="submit" value="ȷ ��"> 
        <input  onClick="javascript:window.document.location.href='xx.asp'" value="�� ��" type=button name="back"> </p>  
 </td>     
    </tr>     
    <tr>      
      <td colspan="4" height="29" width="389">ʦ��������:���书��һ���޸���Ҫ�������书����,����ѡ����ͼƬ.</td>   
    </tr>   
  </table>   
</form>   
<SCRIPT language="JavaScript">    
<!--    
var targetwin="add";   
var badword = new Array("�侫","��","��","ȥ��","��ʺ","��","��","����","��","����","������","����","��","ɵ","��","��","�","����","����","��","����","����","����","����","����","����","غ��","��Ƥ","��ͷ","��","�P","��","�H","����","��","��","����","����","����","����","����","��Ů","����","ү","��","�Ҷ�","��","���","���","��","����","����","�ڹ�","��","ʺ","ƨ","����","�׳�","��","���","���˹�","fuck","cao","��","����");   
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
if(a=="" || a==null){window.alert("����Ҫ��ʽ����");return false;}    
if( IsBadWord(a) ){   
window.alert("��ʽ��������!");   
return false;   
}   
if( CheckIfEnglish(a) ){   
window.alert("��ʹ�ú�������!");   
return false;   
}   
if( IsBadWord(f) ){   
window.alert("��ʽ˵��������!");   
return false;   
}   
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

