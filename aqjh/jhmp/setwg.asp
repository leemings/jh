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
rs.open "SELECT * FROM 用户 where 姓名='"&aqjh_name&"'",conn
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
pai=rs("门派")
rs.close
%>
<html>
<head>
<title><%=pai%>武功设置</title>
<LINK href="../css.css" rel=stylesheet>
</head>
<body background="../jhimg/bk_hc3w.gif">
<div align="center"><font color="#009900"><b>[ <%=pai%> ] 武 功 设 置</b></font><br>
  <font color="#000000" size="2"><a href="addwg.asp"><b>增加招式修改</b></a></font><br>
</div>
<table cellpadding="1" cellspacing="0" border="1" align="center" width="74%" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr> 
    <td height="25" width="107"> 
      <div align="center"> 招 式 名 </div>
    </td>
    <td height="25" width="130"> 
      <div align="center">所 用 武 功</div>
    </td>
    <td height="25" width="141"> 
      <div align="center"> 所 用 内 力 </div>
    </td>
    <td height="25" width="97"> 
      <div align="center">等 级</div>
    </td>
    <td height="25" width="68"> 
      <div align="center"> 操作 </div>
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
        <div align="center"> <font color="#000000" size="2"><a href="modiwg.asp?id=<%=rs("id")%>">修改</a></font> 
          <font color="#000000" size="2"><a href="del.asp?id=<%=rs("id")%>">删除</a></font> 
        </div>
      </td>
    </tr>
    <tr> 
      
    <td height="5" colspan="5"> <font color="#0000FF">说明：<%=rs("f")%><br>
    <%if rs("g")<>"随机" then%>
      <img src="gif/<%=rs("g")%>">
      <%else%>随机图片<%end if%> </font></td>
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
      <div align="center">门派名字：</div>
    </td>
    <td width="76%" height="30"> 
      <input type="text" name="a"  readonly size="15" maxlength="10" value="<%=rs("a")%>">
    </td>
  </tr>
  <tr> 
    <td width="24%"> 
      <div align="center">限制说明：</div>
    </td>
    <td width="76%"> 
      <input type="text" name="e" size="40" value="<%=rs("e")%>" maxlength="45">
      (按下面条件输入<font color="#FFFF00"></font>)</td>
  </tr>
  <tr> 
    <td width="24%"> 
      <div align="center">加入限制：</div>
    </td>
    <td width="76%"> 
      <input type="text" name="g" size="40" value="<%=rs("g")%>" maxlength="45">
      (任意加入填：<font color="#FFFF00">True</font>)</td>
  </tr>
  <tr> 
    <td width="24%"> 
      <div align="center">门派简介：</div>
    </td>
    <td width="76%"> 
        <input type="text" name="d" size="40" value="<%=rs("d")%>" maxlength="45">
      (向大家介绍) </td>
  </tr>
  <tr>
    <td colspan="2" height="15"> 
      <div align="center">
        <input type=submit value=修改 name="submit" style="border: 1px solid; font-size: 9pt; border-color:
#000000 solid">
      </div>
    </td>
  </tr>
  <tr> 
    <td colspan="2" height="30"> 说明：限制说明为加入门派的条件说明，此与表达式相对应的。<br>
      如：想让战斗等级在3级以上的加入,表达式为：<font color="#FFFF00">等级&gt;=3<br>
      </font>如：想让虞斗等级大于2级的女性玩家加入，表达式为：<font color="#FFFF00">等级&gt;=2 and 性别='女'</font><br>
      如：只有已婚玩家，道德大于1000的才可加入，表达式为：<font color="#FFFF00">配偶&lt;&gt;'无' and 道德&gt;1000</font><br>
      如：体力大于1万或内力大于1万的，表达式为：<font color="#FFFF00">体力&gt;10000 or 内力&gt;10000</font><br>
      <font color="#00FF00">如：任意加入，请输入：<b>True</b>(注：以上字母数字全为半角)</font><br>
      <b>以上要按要求去设置，如果设置不下确会不能加入门派。</b></td>
  </tr>
</table>
  </form>
<%rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<p class="p1" align="center">杀伤力=<b><font color="#FF0000">(</font>(</b><font color="#0000FF">自己的攻击-对方的防御</font><b>)</b>+招式的内力+招式的武功<b><font color="#FF0000">)</font></b>x招试等级<br>
  <br>
  如果武功名字不文雅，官府将立刻删除对应帮派，永不恢复！<br>
  <br>
  武功、内力为杀杀力的主要因素，可以根据实际情况进行选择，最好为由小到大递增！<br>
  <br>≮爱情江湖总站≯</p>
</body>

</html>
<SCRIPT language="JavaScript"> 
<!-- 
var targetwin="add";
var badword = new Array("射精","奸","去死","吃屎","妈","娘","日你","尻","操你","干死你","王八","逼","贱","狗","婊","表子","靠你","叉","插你","插死","干你","干死","日死","鸡巴","睾丸","包皮","龟头","屄","赑","妣","肏","奶子","尻","屌","作爱","做爱","床上","抱抱","鸡八","处女","打炮","爷","爸","我儿","·","麻痹","吗比","叼","草你","老吗","乌龟","猪","屎","屁","笨蛋","白痴","滚","大便","死八公","fuck","cao","腚","你吗");
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
var e=document.setmp.e.value;
var g=document.setmp.g.value;
var d=document.setmp.d.value;
if(e=="" || e==null){window.alert("输入要说明文字");return false;}
if(g=="" || g==null){window.alert("输入要条件文字");return false;}
if(d=="" || d==null){window.alert("输入要简介文字");return false;}
if( IsBadWord(e) ){
window.alert("说明不文雅!");
return false;
}
if( IsBadWord(g) ){
window.alert("条件不文雅!");
return false;
}
if( IsBadWord(d) ){
window.alert("简介不文雅!");
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