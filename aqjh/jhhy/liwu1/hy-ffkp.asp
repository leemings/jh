<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
randomize timer
regjm=int(rnd*9998)+1
%>
<html>
<head>
<body bgcolor="#FF9900" background="bg.gif">
<title><%=Application("sjjh_chatroomname")%>领取节日礼物</title>
<link rel="stylesheet" href="../../css.css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#000000" topmargin="2">
<br>
<br>
<center>
  <table width=401 height="227" border=3 align=center cellpadding="5" cellspacing="10" bordercolor="#6633CC" >
      <td width="367" height="252" align="center" valign="middle">输入江湖名称和密码可以领取会员卡片！ 
        <br> <br>
        <table width="325" height="175" align="center">
          <tr>
            <td height="169"> 
              <form method=POST action='hy-ffkp1.asp' onsubmit='return(check());' name=reg>
                <p align="center">
                  <input type=hidden name=regjm value="<%=regjm%>">
                  <br>
                  姓名： 
<input type=text name=name size=12 maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                  <br><br>
                  密码： 
                  <input type=password name=psw size=12 maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                  <br><br>
                  认证： 
                  <input type=password name=psw2 size=4 maxlength="4" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                  请输入：<font color="#FFFF00"><%=regjm%></font></p>
<p align="center">
                  <input type=submit value=领取卡片 name="submit" style="border: 1px solid; font-size: 9pt; border-color:#000000 solid">
				  <input type=button value=关闭 onClick="window.close()" name="button" style="border: 1px solid; font-size: 9pt; border-color:#000000 solid">
</p>
</form>
</td>
</tr>
</table><br>
每逢节日发放一次<br>
        本次只发放[<font color="#FF0000">福神卡</font>],<font color="#FF0000">会员等级</font>张<br>
        下次只发放[<font color="#FF0000">涨钱卡</font>],<font color="#FF0000">会员等级</font>张<br>
        下次只发放[<font color="#FF0000">健身卡</font>],<font color="#FF0000">会员等级</font>张<br>
        时间为：法定节假日<br>
        2015.03.08<br>
        2015.05.01<br>
        2015.06.01<br>
        2015.07.01<br>
        2015.08.01<br>
        2015.完善1<br>
               ……<br>
        [快乐江湖]欢迎大家<a href="../chat/hy.ASP" ><font color="#FFFFFF">购买会员</font></a>！ 
      </td>
    </table>

</center>
</body>
</html>
<SCRIPT language="JavaScript"> 
<!-- 
var targetwin="jhwindow";
function check()
{
var youname=document.reg.name.value;
var mima1=document.reg.psw.value;
var rz=document.reg.psw2.value
var rz1=document.reg.regjm.value
if(youname=="" || youname==null){window.alert("你不报姓名，我怎么知道该给谁呀！！^_^!");return false;}
if( CheckIfEnglish(youname) )
{
window.alert("名字中不能有非法字符、英文、数字；只能使用中文片假名！^_^!");
return false;
}
if(mima1=="" || mima1==null){window.alert("不输入密码可是不行！^_^!");return false;}
if(rz1!=rz){window.alert("请输入正确的认证码！");return false;}
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