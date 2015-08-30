<%Response.Expires=0
Response.Buffer=true
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM 用户 where 姓名='"&sjjh_name&"'",conn,2,2
if rs.bof or rs.eof then
	rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Redirect "error.asp?id=453"
	response.end
end if
if rs("身份")<>"掌门" then
	rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Write "<script language=JavaScript>{alert('不是掌门，不要乱闯！');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
mp=rs("门派")
rs.close
set rs=nothing
conn.close
set conn=nothing
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="setup.css">
<title>武功添加程序♀wWw.happyjh.com♀</title></head>
<body text="#000000" background="../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"><font color="#FFFFFF" size="2"><%=mp%></font>武功添加程序</div>
<form method="post" action="addwgok.asp" name="setmp" onsubmit="return(check());">
  <table border="1" cellspacing="1" align="center" cellpadding="0" bordercolor="#000000" bgcolor="#006699" width="70%">
    <tr> 
      <td width="67"> 
        <div align="center"><font color="#FFFFFF" size="2">门派</font></div>
      </td>
      <td width="197"><font color="#FFFFFF" size="2"> <%=mp%> </font> </td>
      <td width="81"> 
        <div align="center"><font size="2" color="#FFFFFF">武功名</font></div>
      </td>
      <td width="183"><font color="#FFFFFF" size="2"> 
        <input type="text" name="a"
value="" size="15" maxlength="20">
        </font></td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center"><font color="#FFFFFF" size="2">使用武功</font></div>
      </td>
      <td width="197"><font color="#FFFFFF" size="2"> 
        <input type="text" name="c"
value="1" size="15" maxlength="3">
        </font></td>
      <td width="81" nowrap> 
        <div align="center"><font color="#FFFFFF" size="2">使用内力</font></div>
      </td>
      <td width="183"><font color="#FFFFFF" size="2"> 
        <input type="text" name="d"
value="1" size="15" maxlength="3">
        </font></td>
    </tr>
    <tr> 
      <td width="67" nowrap> 
        <div align="center"><font size="2" color="#FFFFFF">发招说明</font></div>
      </td>
      <td colspan="3"><font color="#FFFFFF" size="2"> 
        <input type="text" name="f" size="50" maxlength="50">
        </font></td>
    </tr>
    <tr> 
      <td colspan="4"> 
        <input type="submit" value="确 定">
        <input  onClick="javascript:window.document.location.href='setwg.asp'" value="返 回" type=button name="back">
        <font color="#FFFFFF" size="-1">招式图片： 
        <select name=g size=1 onChange="document.images['face'].src='gif/'+options[selectedIndex].value+'';" style="BACKGROUND-COLOR: #99CCFF; BORDER-BOTTOM: 1px double; BORDER-LEFT: 1px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000">
          <option value="随机" selected>随机</option>
          <%for t=1 to 24%>
          <option value="wg<%=t%>.gif">wg<%=t%>.gif</option>
			<%next%>
		 </select>
        <img id=face src="gif/wg1.gif" alt=招式图片></font> </td>
    </tr>
    <tr> 
      <td colspan="4" height="29"> <font color="#FFFFFF" size="-1">说明:发招说明为在使用此招式时的描述,其中##为发招人,%%代表被打人。<br>
        招式名：<font color="#FFFF00">烈火金刚</font> <br>
        说明为：<font color="#FFFF00">##吸气，呼气。只见双掌发红…如烈火般燃烧，重重的向%%拍去……</font></font></td>
    </tr>
  </table>
</form>
<SCRIPT language="JavaScript"> 
<!-- 
var targetwin="add";
var badword = new Array("射精","奸","去死","吃屎","妈","娘","日你","尻","操你","干死你","王八","逼","傻","贱","狗","婊","表子","靠你","叉","插你","插死","干你","干死","日死","鸡巴","睾丸","包皮","龟头","屄","赑","妣","肏","奶子","尻","屌","作爱","做爱","床上","抱抱","鸡八","处女","打炮","爷","爸","我儿","·","麻痹","吗比","叼","草你","老吗","乌龟","猪","屎","屁","笨蛋","白痴","滚","大便","死八公","fuck","cao","腚","你吗");
var badstr = "~!@ #$%^&*()[]{}_+-|=\`;,:'\"?<>/～！·＃￥％…；‘’：“”—＊（　）—＋｜－＝、／。，？《》" ;
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
if(a=="" || a==null){window.alert("输入要招式名字");return false;}
if(f=="" || f==null){window.alert("输入要说明文字");return false;}
if( IsBadWord(a) ){
window.alert("招式名不文雅!");
return false;
}
if( CheckIfEnglish(a) ){
window.alert("请使用汉字输入!");
return false;
}
if( IsBadWord(f) ){
window.alert("招式说明不文雅!");
return false;
}
}
function check1()
{
var youname=document.modi.wgname.value;
if(youname=="" || youname==null){window.alert("请输入武功名字！！^_^!");return false;}
if( IsBadWord(youname) ){
window.alert("招式名字不文雅!");
return false;
}
if( CheckIfEnglish(youname) )
{
window.alert("武功中不能有非法字符、英文、数字；只能使用中文片假名！^_^!");
return false;
}
return true;
}

function CheckIfEnglish( String )
{ 
    var Letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890`=~!@#$%^&*()_+[]{}\\|/?.>,<;:'\-?<>/～！·＃￥％…；‘’：“”—＊（　）—＋｜－＝、／。，？《》↑↓⊙●★☆■♀ 『』◆◥◤◣ Ψ ※ →№←㊣∑⌒ 〖〗 ＠ξζω□ ∮〓※ ▓∏卐【 】▲△√ ∩¤々 ♀♂∞ ㄨ ≡↘↙ ＆◎Ю┼┏ ┓田 ┃▎○╪┗┛ ∴ ①②③④⑤⑥⑦⑧ \"";
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