<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
id=LCase(trim(request.querystring("id")))
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM 用户 where 姓名='"&sjjh_name&"'",conn
if rs.bof or rs.eof then
	rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Redirect "error.asp?id=453"
	response.end
end if

pai=sjjh_name
rs.close
%>
<html>
<head>
<title><%=pai%>轩辕武功设置♀wWw.51eline.com♀</title>
<style type="text/css">
td           { font-family: 宋体; font-size: 12px }
body         { font-family: 宋体; font-size: 12px;
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff}
</style>
<LINK href="../css.css" rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#006699" leftmargin="0" topmargin="2">
<div align="center"><font color="#009900"><b><font color="#FFFFFF" size="3">[ </font><font color="#ff0000" size="3"><%=pai%></font>
  <font color="#FFFFFF" size="3">]</font><font color="#FFFFFF"> <font size="3">轩 
  辕 武 功 设 置</font></font></b></font><br>
  <br>
  <a href="addwg.asp"><font color="#FF6600" size="2">增加轩辕招式</font></a><br>
</div><br>
<table width="80%" border="1" align="center" cellpadding="1" cellspacing="0" bordercolorlight="#006699" bordercolordark="#efefef" bgcolor="#9999CC">
  <tr> 
    <td height="25" width="107"> 
      <div align="center"> <font size="2">招 式 名</font> </div>
    </td>
    <td height="25" width="130"> 
      <div align="center"><font size="2">所 用 武 功</font></div>
    </td>
    <td height="25" width="141"> 
      <div align="center"> <font size="2">所 用 内 力</font> </div>
    </td>
    <td height="25" width="97"> 
      <div align="center"><font color="#000000" size="2">等 级</font></div>
    </td>
    <td height="25" width="130"> 
      <div align="center"><font size="2">修炼次数</font></div>
    </td>	
    <td height="25" width="150"> 
      <div align="center"> <font size="2">操作</font> </div>
    </td>
  </tr>
  <%
rs.open "SELECT * FROM n where b='" & pai & "' order by (c+d)",conn
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
      <div align="center"><font color="#000000"> <%=int(sqr((rs("e")/20)))+1%> 
          </font></div>
      </td>
    <td height="2" width="130"> 
      <div align="center"> <font color="#000000"><%=rs("e")%></font></div>
      </td>
      <td height="2" width="150"> 
        <div align="center">          <font color="#000000" size="2"><a href="xl.asp?id=<%=rs("id")%>">修炼</a></font> <font color="#000000" size="2"><a href="modiwg.asp?id=<%=rs("id")%>">修改</a></font> 
	  <font color="#000000" size="2"><a href="del.asp?id=<%=rs("id")%>">删除</a></font>
        </div>
      </td>
    </tr>
    <tr> 
      
    <td height="5" colspan="9"> <font color="#0000FF">说明：<%=rs("f")%><br>
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
<table width="80%" border="1" align="center" cellpadding="1" cellspacing="0" bordercolorlight="#006699" bordercolordark="#efefef" bgcolor="#336699">
  <tr>
    <td><font color="#FFFFFF" size="2">杀伤力=</font><font size="2"><b><font color="#FF0000">(</font><font color="#FFFFFF">(</font></b><font color="#FF0000">自己的攻击-对方的防御</font><b><font color="#FFFFFF">)</font></b><font color="#FFFFFF">+招式的内力+招式的武功</font><b><font color="#FF0000">)</font></b><font color="#FFFFFF">x招式等级</font></font></td>
  </tr>
  <tr>
    <td><font color="#FFFFFF" size="2">如果武功名字不文雅，官府将立刻删除对应用户，永不恢复！<br>
      武功、内力为杀杀力的主要因素，可以根据实际情况进行选择，最好为由小到大递增！</font></td>
  </tr>
  <tr>
    <td><font color="#990000" size="2">轩辕等级需要靠修炼来提高，提升1级至少修炼20次！</font></td>
  </tr>
  <tr>
    <td><font color="#990000" size="2">每次修炼需花费金、木、水、火、土及银币各60！</font></td>
  </tr>
  <tr>
    <td><font color="#990000" size="2">每级轩辕的特效为:1级吸法力;2-3级吸轻功;4-5级吸防御;6-7级吸攻击;8-9级吸银币;10级及以上吸金币</font></td>
  </tr>
</table>
<div align="center"><br>
  <%
rs.open "SELECT * FROM p where a='" & pai & "'",conn
%>
  <%rs.close 
set rs=nothing 
conn.close 
set conn=nothing 
%>
  <br>
  <font color="#FFFFFF">【 授权：<%=Application("sjjh_user")%> </font><font color="#FFFFFF">运行序列号：<%=Application("sjjh_sn")%></font><font color="#FFFFFF">】</font></p> 
</div>
</body>  
  
</html>  
<SCRIPT language="JavaScript">   
<!--   
var targetwin="add";  
var badword = new Array("江湖总站","世纪江湖","天外天江湖","六六江湖","学子天地","破江湖","玩别的江湖","去别的江湖","去我的江湖","来我的江湖","玩我的江湖","去其它江湖","去其他江湖","玩其他江湖","玩其它江湖","烂江湖","滥江湖","射精","奸","我干","鸟人","发全","喀什","咔什","cfq","爬你达来蛋","撅你达来蛋","死你达来蛋","达来蛋","chenfaquan","chen fa quan","chen","ma de","kan ni","fa","quan","ni ma","mao","ni mu","niba","ni ba","nima","nimu","ni nai nai","ni jie","ni mei","nai","ye","ma","ba","zu","zhu","傻鸟","站掌","占长","战长","丈长","帐长","我塞","我靠","老师","砍你","砍了你","砍死你","1线","捅你","捅死你","教师","妓女","陈发","做鸭","作鸭","紫华","sai","管理员","网管","kao","cao","yixiantian","yi xian tian","yxt","eline","塞你母","8196316","13808556346","一线","一线天","线天","一钱天","你母","你木","发铨","去死","站长","老大","吃屎","妈","爸","爹","娘","日你","尻","爷爷","奶奶","你爷","你奶","他爷","他奶","操你","我操","干死你","王八","逼","傻","贱","剁了你","做婊","做表","作婊","作表","婊","表子","靠你","叉","插你","插死","干你","干死","日死","鸡巴","睾丸","包皮","龟头","屄","赑","妣","肏","奶子","尻","屌","作爱","做爱","床上","抱抱","鸡八","处女","处男","打炮","爸","我儿","麻痹","吗比","叼","草你","阿冒","阿茂","啊冒","啊茂","老吗","乌龟","屎","屁","笨蛋","白痴","滚","江湖人","大便","死八公","http","www","com","cn","org","net","://","fuck","cao","gan","sai","□","你吗");  
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