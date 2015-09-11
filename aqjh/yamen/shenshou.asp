<%@ LANGUAGE=VBScript codepage ="936" %>
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")

randomize timer
regjm=int(rnd*8998)+1000
id=LCase(trim(request.querystring("id")))
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');window.close();}</script>"
Response.End 
end if

if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 召唤兽1,转生 from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
zaohuan1=rs("召唤兽1")
zhuan=rs("转生")
if zaohuan1<>"无" then
Response.Write "<script language=JavaScript>{alert('你有神兽了');window.close();}</script>"
    response.end

	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if
if zhuan<8 then
Response.Write "<script language=JavaScript>{alert('至少需要八次转生以上，你有那资格么?');window.close();}</script>"
    response.end

	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if

set rs=nothing	
conn.close
set conn=nothing
%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<title><%=Application("aqjh_chatroomname")%>神兽注册</title>
<link rel="stylesheet" href="../css.css">
</head>
<body bgcolor="#800000" background="../chat/bg.gif">
<center>
<table border=1 bgcolor="#FFFFFF" align=center width=478 cellpadding="5" cellspacing="10" background="../images/b2.gif" height="438">
<tr bgcolor="#FFFFFF" align="center">
      <td height="525" bgcolor="0088FF" width="444"> <font color="#FF6600">请大侠仔细填写神兽信息</font> 
        <br>
<br>
<table width="366" height="229">
<tr>
            <td height="275" width="358"> 
              <form method=POST action='shenshouok.asp' onsubmit='return(check());' name=reg>
                <p align="left">神兽姓名：                        
                  <input type=text name=name size=12 maxlength="5" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                  <font color="#FF0000">*</font> 神兽名5个汉字                        
                  <input type=hidden name=regjm value="<%=regjm%>">
                  <br>
                  登录密码：                        
                  <input type=password name=psw size=12 maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="wjgy26z">
                  <font color="#FF0000">*</font>要求6-10个字符，嘿嘿<br>                       
                  确认密码：                        
                  <input type=password name=pswc size=12 maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="wjgy26z">
                  <font color="#FF0000">*</font>同上 <br>                       
                  <br>
                  第二密码：                        
                  <input type=password name=psw2 size=12 maxlength="15" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="aqjhwmhwl">
                  <font color="#FF0000">*</font>要求10-15个字符<br>                       
                  确认密码：                        
                  <input type=password name=pswc2 size=12 maxlength="15" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="aqjhwmhwl">
                  <font color="#FF0000">*</font>同上 <br>非常重要，用于ID被盗时重新设置登录密码<br>                       
                  性别：                        
                  <select name=sex size=1 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                    <option value=男 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'"selected>男 
                    <option value=女 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">女 
                  </select>
                  <br>
                  OICQ：                        
                  <input type="text" name="oicq" size="11" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" maxlength="11" value="0000000">
                  &nbsp; <font color="#FF0000">*</font>填写QQ好与朋友联系5位以上！ <br>                       
                  信箱：                        
                  <input type=text name=e_mail size=25 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="cys0831@163.com" maxlength="30">
                  <font color="#FF0000">*</font>联系使用<br>                       
                  介绍人：                                      
                  <input size=10 name=jsr value="永不放弃"maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">                  <br>
 <br>
                  名单头像：<img id=face src="../ico/n.gif" alt=个人形象代表 width="14" height="16"> 
<select name=face size=1 onChange="document.images['face'].src='../ico/'+options[selectedIndex].value+'-2.gif';">         
                    <%for i=1 to 108%>
                    <option value='<%=i%>'><%=i%></option>
                    <%next%>
                  </select>                  头像将在聊天室中显示！</p>                       
                  <input type=hidden name=regjm1 value="<%=regjm%>">认证码：<input type=text name=regjm2 size=4 maxlength="4" onKeypress="if (event.keyCode < 45 || event.keyCode > 57) {event.returnValue = false;}"  style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="<%=regjm%>"> 请在左侧输入右边的数字<font color="#ff0000"> <%=regjm%> </font>                       
				
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
var mima2=document.reg.pswc.value;
var mima3=document.reg.psw2.value;
var mima4=document.reg.pswc2.value;
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
if(mima3=="" || mima3==null){window.alert("第二密码很重要，一定要有！当ID被盗时可以用于重新设置登录密码！^_^!");return false;}
if(mima4=="" || mima4==null){window.alert("验证第二密码不对！^_^!");return false;}
if(mima3!=mima4){window.alert("两次输入密码不相同！");return false;}
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
