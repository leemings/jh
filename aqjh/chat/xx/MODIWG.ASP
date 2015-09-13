<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
id=clng(Request("id"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sqlstr="select * from n where id="&id
rs.open sqlstr,conn
if rs.EOF or rs.BOF then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：没有你选择的武功。');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style type="text/css">
td           { font-family: 宋体; font-size: 12px }
body         { font-family: 宋体; font-size: 12px;}
</style>
<title>轩辕武功修改程序</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#006699" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF" leftmargin="0" topmargin="2">
<div align="center"><font color="#FFFFFF"><b>轩辕武功修改程序</b></font></div>
<form method="post" action="modiwgok.asp" name="setmp" onsubmit="return(check());">
  <table border="1" cellspacing="1" align="center" cellpadding="0" bordercolor="#000000" bgcolor="#9999CC" width="70%">
    <tr> 
      <td width="67"> 
        <div align="center"><font color="#000000">ID</font></div>
      </td>
      <td width="197"><font color="#FFFFFF"> <%=rs("id")%> </font> 
        <input type="hidden" name="id" value="<%=rs("id")%>" size="15" maxlength="20">
        <font color="#000000">等级:</font><font color="#FFFFFF"><%=int(sqr((rs("e")/100)))+1%></font> 
      </td>
      <td width="81"> 
        <div align="center"><font color="#000000">武功名</font></div>
      </td>
      <td width="183"><font color="#FFFFFF"> 
        <input type="text" name="a"
value="<%=rs("a")%>" size="15" maxlength="20">
        </font></td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center"><font color="#000000">使用武功</font></div>
      </td>
      <td width="197"><font color="#FFFFFF"> 
        <input type="text" name="c"
value="<%=rs("c")%>" size="15" maxlength="3">
        </font></td>
      <td width="81" nowrap> 
        <div align="center"><font color="#000000">使用内力</font></div>
      </td>
      <td width="183"><font color="#FFFFFF"> 
        <input type="text" name="d"
value="<%=rs("d")%>" size="15" maxlength="3">
        </font></td>
    </tr>
    <tr> 
      <td width="67" nowrap> 
        <div align="center"><font color="#000000">发招说明</font></div>
      </td>
      <td colspan="3"><font color="#FFFFFF"> 
        <input type="text" name="f"
value="<%=rs("f")%>" size="50" maxlength="50">
        </font></td>
    </tr>
    <tr> 
      <td colspan="4"> 
        <div align="center"></div>
        <input type="submit" value="确 定" name="submit">
        <input  onClick="javascript:window.document.location.href='setwg.asp'" value="返 回" type=button name="back">
        <font color="#000000">招式图片：</font><font color="#FFFFFF"> 
        <select name=g size=1 onChange="document.images['face'].src='gif/'+options[selectedIndex].value+'';" style="BACKGROUND-COLOR: #99CCFF; BORDER-BOTTOM: 1px double; BORDER-LEFT: 1px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000">
          <option value="随机" selected>随机</option>
          <%for t=1 to 24%>
          <option value="wg<%=t%>.gif">wg<%=t%>.gif</option>
          <%next%>
        </select> 
        <img id=face src="gif/wg1.gif" alt=招式图片></font> </td> 
    </tr>
    <tr> 
      <td colspan="4" height="29"> <font color="#FF0000">说明:</font><font color="#000000">发招说明为在使用此招式时的描述,其中##为发招人,%%代表被打人。</font><font color="#FFFFFF"><br>
        <font color="#990000">招式名：</font><font color="#FFFF00">烈火金刚</font> <br>
        <font color="#990000">说明为：</font><font color="#FFFF00">##吸气，呼气。只见双掌发红…如烈火般燃烧，重重的向%%拍去……</font></font></td>
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
var badword = new Array("江湖总站","快乐江湖","天外天江湖","六六江湖","学子天地","破江湖","玩别的江湖","去别的江湖","去我的江湖","来我的江湖","玩我的江湖","去其它江湖","去其他江湖","玩其他江湖","玩其它江湖","烂江湖","滥江湖","射精","奸","我干","鸟人","发全","喀什","咔什","cfq","爬你达来蛋","撅你达来蛋","死你达来蛋","达来蛋","chenfaquan","chen fa quan","chen","ma de","kan ni","fa","quan","ni ma","mao","ni mu","niba","ni ba","nima","nimu","ni nai nai","ni jie","ni mei","nai","ye","ma","ba","zu","zhu","傻鸟","站掌","占长","战长","丈长","帐长","我塞","我靠","老师","砍你","砍了你","砍死你","1线","捅你","捅死你","教师","妓女","陈发","做鸭","作鸭","紫华","sai","管理员","网管","kao","cao","yixiantian","yi xian tian","yxt","eline","塞你母","8196316","13808556346","一线","一线天","线天","一钱天","你母","你木","发铨","去死","站长","老大","吃屎","妈","爸","爹","娘","日你","尻","爷爷","奶奶","你爷","你奶","他爷","他奶","操你","我操","干死你","王八","逼","傻","贱","剁了你","做婊","做表","作婊","作表","婊","表子","靠你","叉","插你","插死","干你","干死","日死","鸡巴","睾丸","包皮","龟头","屄","赑","妣","肏","奶子","尻","屌","作爱","做爱","床上","抱抱","鸡八","处女","处男","打炮","爸","我儿","麻痹","吗比","叼","草你","阿冒","阿茂","啊冒","啊茂","老吗","乌龟","屎","屁","笨蛋","白痴","滚","江湖人","大便","死八公","http","www","com","cn","org","net","://","fuck","cao","gan","sai","□","你吗");
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
