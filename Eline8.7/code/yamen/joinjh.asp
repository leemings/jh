<%@ LANGUAGE=VBScript codepage ="936" %>
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
randomize timer
regjm=int(rnd*8998)+1000
id=LCase(trim(request.querystring("id")))
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');window.close();}</script>"
Response.End 
end if
%>
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>
<title><%=Application("sjjh_chatroomname")%>用户注册</title>
<link rel="stylesheet" href="../css.css">
</head>
<body bgcolor="#000000" leftmargin="0" topmargin="0">
<center>
<script language="JavaScript">
if (navigator.appName.indexOf("Internet Explorer") != -1)	
document.onmousedown = noright;
function noright()
{
if (event.button == 2 | event.button == 3)
{
	alert("不要这样嘛 ^_^");
	 location.replace("#")
}
}
</script>
  <table border=1 bgcolor="#000000" align=center width=400 cellpadding="3" cellspacing="10" height="500">
    <tr bgcolor="#FFFFFF" align="center">
      <td height="480" bgcolor="#000000"><font color="#FF6600">请大侠仔细填写以下信息</font>
<table width="325" height="229">
<tr>
            <td height="275"> 
              <form method=POST action='joinjhnow.asp' onsubmit='return(check());' name=reg>
                <p align="left"><font color="#FFFFFF">姓名：</font> 
                  <input type=text name=name size=12 maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #ffffff" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                  <font color="#FF0000">*</font> <font color="#FFFFFF">用户名5个汉字</font> 
                  <input type=hidden name=regjm value="<%=regjm%>">
                  <br>
                  <font color="#FFFFFF">密码：</font> 
                  <input type=password name=psw size=12 maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #ffffff" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                  <font color="#FF0000">*</font><font color="#FFFFFF">最少6-10个字符</font><br>
                  <font color="#FFFFFF">确认：</font> 
                  <input type=password name=pswc size=12 maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #ffffff" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                  <font color="#FF0000">*</font><font color="#FFFFFF">同上 </font><br>
                  <font color="#FFFFFF">性别：</font> 
                  <select name=sex size=1 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #ffffff" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                    <option value=男 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #ffffff" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'"selected>男 
                    <option value=女 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #ffffff" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">女 
                  </select>
                  <br>
                  <font color="#FFFFFF">Q  Q：</font> 
                  <input type="text" name="oicq" size="11" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #ffffff" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" maxlength="11">
                  <font color="#FF0000">*</font><font color="#FFFFFF">填写QQ好与朋友联系5位以上！</font> 
                  <br>
                  <font color="#FFFFFF">信箱：</font> 
                  <input type=text name=e_mail size=25 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #ffffff" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="@" maxlength="30">
                  <font color="#FF0000">*</font><font color="#FFFFFF">联系使用</font><br>
                  <font color="#FFFFFF">介绍人：</font> 
                  <input size=10 name=jsr value="<%=id%>"maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #ffffff" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                  <br>
                  <font color="#FFFFFF">[此项介绍姓名，只为系统记录，没有可以不填写]</font><br>
                  <font color="#FFFFFF">名单头像：</font><img id=face src="../ico/n.gif" alt=个人形象代表> 
                  <select name=face size=1 onChange="document.images['face'].src='../ico/'+options[selectedIndex].value+'-2.gif';" style="BACKGROUND-COLOR: #99CCFF; BORDER-BOTTOM: 1px double; BORDER-LEFT: 1px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000">
                    <option selected value="00">请选择</option>
                    <%for i=1 to 468%>
                    <option value='<%=i%>'><%=i%></option>
                    <%next%>
                  </select> <font color="#FF0000">*</font>
                  <font color="#FFFFFF">头像将在聊天室中显示！</font></p>
				  <input type=hidden name=regjm1 value="<%=regjm%>">
                <font color="#FF0000">认证码：</font>：<input type=text name=regjm2 size=4 maxlength="4" onKeypress="if (event.keyCode < 45 || event.keyCode > 57) {event.returnValue = false;}"  style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #ffffff" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                <font color="#FFFFFF">请在左侧输入右边的数字</font><font color="#009933"> 
                <%=regjm%></font> 
                <p align="center">
<input type=submit value=申请 name="submit" style="border: 1px solid; font-size: 9pt; border-color:
#000000 solid">
<input type=button value=关闭 onClick="window.close()" name="button" style="border: 1px solid; font-size: 9pt; border-color:
#000000 solid">
</p>
</form>
</td>
</tr>
</table>
        <div align="center"><img src="1.gif" alt="准备已就绪∥创荡E线去" width="100" height="100"> 
        </div>
        <div align="center">
          <table width="325" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><font color="#CC0000">请注意</font><font color="#FFFFFF">：</font><font color="#CCCCCC">当你踏足本江湖后.也就踏上了你的江湖生涯.在这里只能靠自己.没有人会帮你.如果想要站长帮你改数据.那请你最好别注册.因为那是没可能的事！</font></td>
            </tr>
          </table>
        </div></table>

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
var mima2=document.reg.pswc.value;
var e_mail=document.reg.e_mail.value;
var oicq=document.reg.oicq.value;
var regno1=document.reg.regjm1.value;
var regno2=document.reg.regjm2.value;
if(youname=="" || youname==null){window.alert("输入要注册的名字！！^_^!");return false;}
if( CheckIfEnglish(youname) )
{
window.alert("名字中不能有非法字符、英文、数字；只能使用中文片假名！^_^!");
return false;
}
if(mima1=="" || mima1==null){window.alert("不输入密码可是不行！^_^!");return false;}
if(mima2=="" || mima2==null){window.alert("验证密码不对！^_^!");return false;}
if(mima2!=mima2){window.alert("两次输入密码不相同！");return false;}
if(regno1!=regno2){window.alert("请输入正确的认证码:"+regno1+"！");return false;}
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