<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
if sjjh_name<>Application("sjjh_user") then 
	Response.Write "<script Language=Javascript>alert('提示：你不是正站长，你不能操作！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id=clng(Request("wgid"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sqlstr="select * from n where id="&id
rs.open sqlstr,conn
if rs.EOF or rs.BOF then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：没有你选择的武功！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../JHMP/setup.css">
<title>轩辕修改程序♀wWw.happyjh.com♀</title></head>
<body text="#000000" background="../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center">轩辕修改程序</div>
<div align="center">人名=<%=rs("b")%></div>
<form method="post" action="managexxok.asp" name="setmp" onsubmit="return(check());">
  <table border="1" cellspacing="1" align="center" cellpadding="0" bordercolor="#000000" bgcolor="#006699" width="70%">
    <tr> 
      <td width="67"> 
        <div align="center"><font color="#FFFFFF" size="2">ID</font></div>
      </td>
      <td width="197"><font color="#FFFFFF" size="2"> <%=rs("id")%> </font> 
        <input type="hidden" name="id" value="<%=rs("id")%>" size="15" maxlength="20">
        <font color="#FFFFFF" size="2">等级:<%=(int(sqr(rs("e")/100))+1)%></font> </td>
      <td width="81"> 
        <div align="center"><font size="2" color="#FFFFFF">轩辕名</font></div>
      </td>
      <td width="183"><font color="#FFFFFF" size="2"> 
        <input type="text" name="a"
value="<%=rs("a")%>" size="15" maxlength="20">
        </font></td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center"><font color="#FFFFFF" size="2">使用武功</font></div>
      </td>
      <td width="197"><font color="#FFFFFF" size="2"> 
        <input type="text" name="c"
value="<%=rs("c")%>" size="15" maxlength="3">
        </font></td>
      <td width="81" nowrap> 
        <div align="center"><font color="#FFFFFF" size="2">使用内力</font></div>
      </td>
      <td width="183"><font color="#FFFFFF" size="2"> 
        <input type="text" name="d"
value="<%=rs("d")%>" size="15" maxlength="3">
        </font></td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center"><font color="#FFFFFF" size="2">修改等级</font></div>
      </td>
      <td width="197"><font color="#FFFFFF" size="2"> 
        <input type="text" name="e"
value="<%=rs("e")%>" size="15" maxlength="7">
	</font></td>
    </tr>
    <tr> 
      <td width="67" nowrap> 
        <div align="center"><font size="2" color="#FFFFFF">发招说明</font></div>
      </td>
      <td colspan="3"><font color="#FFFFFF" size="2"> 
        <input type="text" name="f"
value="<%=rs("f")%>" size="50" maxlength="50">
        </font></td>
    </tr>
    <tr> 
      <td colspan="4"> 
        <div align="center"></div>
        <input type="submit" value="确 定" name="submit">
        <input  onClick="javascript:window.document.location.href='adminxx.asp?mp=<%=rs("b")%>'" value="返 回" type=button name="back">      
        <font color="#FFFFFF" size="-1">招式图片：       
        <select name=g size=1 onChange="document.images['face'].src='gif/'+options[selectedIndex].value+'';" style="BACKGROUND-COLOR: #99CCFF; BORDER-BOTTOM: 1px double; BORDER-LEFT: 1px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000">      
   <option value="随机" selected>随机</option>
          <%for t=1 to 24%>
          <option value="wg<%=t%>.gif">wg<%=t%>.gif</option>
			<%next%>
	    </select>      
        <img id=face src="../xx/gif/wg1.gif" alt=招式图片></font> </td>      
    </tr>      
    <tr>       
      <td colspan="4" height="29"> <font color="#FFFFFF" size="-1">说明:发招说明为在使用此招式时的描述,其中##为发招人,%%代表被打人。<br>      
        招式名：<font color="#FFFF00">烈火金刚</font> <br>      
        说明为：<font color="#FFFF00">##吸气，呼气。只见双掌发红…如烈火般燃烧，重重的向%%拍去……</font></font></td>      
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