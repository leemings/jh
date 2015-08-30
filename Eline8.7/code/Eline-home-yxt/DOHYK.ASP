<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
%>
<html>
<head>
<title>会员数据库管理♀wWw.happyjh.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>
<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif">
<p align="center"><font color="#0000FF">会员冲卡</font></p>
<p align="center"><font color="#0000FF">&nbsp;</font></p>
<p align="center">　</p>
<table border="1" width="35%" bordercolorlight="#000000" cellspacing="1" cellpadding="1" bordercolordark="#85C2E0" height="141" align="center">
  <tr> 
    <td bgcolor="#85C2E0"   width="51%" align="center" nowrap height="131">
    　<form method="POST" action="dohykok.asp" onsubmit='return(check());' name="loginhyk">
      <p align="center">
      卡号：<input type="text" name="kname" size="20"><br>密码：<input type="text" name="kpass" size="20"><br>
      姓名：<input type="text" name="name" size="20"><br>密码：<input type="text" name="pass" size="20"></p>
      <p align="center">
      <input type="submit" value="确定" name="submit" style="border: 1px solid; font-size: 9pt; border-color:
#000000 solid"></p>
    </form>
    <p>　</td>
  </tr>
  </table>
<p align="center">　</p>
<p align="center">本程序由休闲江湖修改制作</p>
</body>
</html>
<script language="JavaScript">
<!--
<!--
ie = (document.all)? true:false
if (ie){function ctlent(eventobject){if(window.event.keyCode==13 && check()!=false){this.document.login.submit();}}}
function check()
{
var name=document.loginhyk.name.value;
var pass=document.loginhyk.pass.value;
var kname=document.loginhyk.kname.value;
var kpass=document.loginhyk.kpass.value;
if(name=="" || name==null){window.alert("请输入你的姓名");return false;}
if(pass=="" || pass==null){window.alert("请输入你江湖上的密码");return false;}
if(kname=="" || kname==null){window.alert("请输入会员卡号");return false;}
if(kpass=="" || kpass==null){window.alert("请输入会员卡密码");return false;}
if( CheckIfEnglish(name) )
{
window.alert("名字中不能有非法字符、英文、数字；只能使用中文片假名！^_^!");
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
</script >