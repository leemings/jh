<%Response.Expires=0
Response.Buffer=true
un=session("yx8_mhjh_username")
if un="" then Response.Redirect "../../error.asp?id=016"
%><head>
<link rel=stylesheet href='../../style.css'>
<title>武功添加</title></head>
<body background='../../chatroom/bg1.gif'>
<div align="center">添加轩辕武功</div>
<form method="post" action="add2.asp" name="setmp" onsubmit="return(check());">
  <table border="1" cellspacing="1" align="center" cellpadding="0" bordercolor="#000000" width="397">
    <tr> 
      <td width="65"> 
        <div align="center">轩辕大侠</div>
      </td>
      <td width="134"><%=un%></td>
      <td width="67"> 
        <div align="center">武功名</div>
      </td>
      <td width="111"> 
        <input type="text" name="a"
value="" size="14" maxlength="4">
        </td>
    </tr>
    <tr> 
      <td width="389" colspan="4"> 
                            <select name="b">
                            <option value="攻击型" selected>选择武功类型</option>                
                            <option value="gj">攻击型武功</option>
                            <option value="fs">防守型武功</option>
                            </select> 
        <select name=c size=1 onChange="document.images['face'].src='gif/'+options[selectedIndex].value+'';" style="BACKGROUND-COLOR: #99CCFF; BORDER-BOTTOM: 1px double; BORDER-LEFT: 1px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000">    
          <option value="随机" selected>招式图片</option>
          <option value="随机">随机图片</option>
         <%for t=1 to 24%>
          <option value="wg<%=t%>.gif">wg<%=t%>.gif</option>
			<%next%>
		 </select>    
        <img id=face src="gif/wg1.gif" alt=随机图片>  
      </td> 
    </tr> 
    <tr>  
      <td colspan="4" width="389">    
       <p align="center">   
       <input type="submit" value="确 定">
        <input  onClick="javascript:window.document.location.href='xx.asp'" value="返 回" type=button name="back"> </p> 
 </td>    
    </tr>    
    <tr>     
      <td colspan="4" height="29" width="389">说明:请在武功名一栏填写你要创的个人秘技,还有选择类型和攻击图片.</td>   
    </tr>   
  </table>   
</form>   
<SCRIPT language="JavaScript">    
<!--    
var targetwin="add";   
var badword = new Array("射精","奸","靠","去死","吃屎","妈","娘","日你","尻","操你","干死你","王八","逼","傻","贱","狗","婊","表子","靠你","叉","插你","插死","干你","干死","日死","鸡巴","睾丸","包皮","龟头","屄","赑","妣","肏","奶子","尻","屌","作爱","做爱","床上","抱抱","鸡八","处女","打炮","爷","爸","我儿","·","麻痹","吗比","叼","草你","老吗","乌龟","猪","屎","屁","笨蛋","白痴","滚","大便","死八公","fuck","cao","腚","你吗");   
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
if(a=="" || a==null){window.alert("输入要招式名字");return false;}    
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
