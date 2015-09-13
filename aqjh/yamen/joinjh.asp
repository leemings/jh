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
<HTML><HEAD><TITLE>注册帐号</TITLE>
<LINK href="pic/lyy.css" rel=stylesheet>
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
</HEAD>
<BODY background=pic/BG.gif oncontextmenu=self.event.returnValue=false>
<TABLE id=AutoNumber1 style="BORDER-COLLAPSE: collapse" borderColor=#111111 cellSpacing=0 cellPadding=0 width=380 align=center border=0>
<TBODY><TR>
<TD width="3%" height=1><IMG src="pic/T1.gif" border=0></TD>
<TD width="94%" background=pic/M1.gif height=1>　</TD>
<TD width="12%" height=1><IMG src="pic/T2.gif" border=0></TD></TR>
  <TR><TD width="3%" background=pic/M2.gif rowSpan=3></TD>
    <TD width="94%" height=23><P align=center><SPAN lang=en>&copy;   
<B>请大侠仔细填写以下信息</P></SPAN></b></TD>
    <TD width="12%" background=pic/M2.gif height=1 rowSpan=3></TD></TR>
  <TR>
    <TD width="94%" height=16>
      <HR color=#682420 SIZE=1>
    </TD></TR>
  <TR><TD width="94%">
<!=><form method=POST action='joinjhnow.asp' onsubmit='return(check());' name=reg>
<TABLE height=403 cellSpacing=1 width="99%" border=0>
        <TBODY>
        <TR>
          <TD width="21%" height=25>*姓名：</TD>
          <TD width="31%" height=25><INPUT class=input maxLength=10 size=11 type=text name=name></TD>
          <TD width="48%" height=25>必须使用2-5个汉字</TD></TR>
        <TR>
          <TD width="21%" height=25>*性别：</TD>
          <TD width="31%" height=25><SELECT size=1 name=sex> 
<OPTION value=男 selected>男</OPTION><OPTION value=女>女</OPTION></SELECT></TD>
          <TD width="48%" height=25>请选择你的性别</TD></TR>
        <TR>
          <TD width="21%" height=25>*头像：</TD>
          <TD width="31%" height=25><select name=face size=1 onChange="document.images['face'].src='../ico/'+options[selectedIndex].value+'-2.gif';">
                    <%for i=1 to 108%>
                    <option value='<%=i%>'><%=i%></option>
                    <%next%>
                  </select></TD>
          <TD width="48%" height=25><IMG id=face src="../ico/n.gif"></TD></TR>
        <TR>
          <TD width="21%" height=25>*密码：</TD>
          <TD width="31%" height=25><INPUT class=input type=password 
            maxLength=10 size=11 name=psw></TD>
          <TD width="48%" height=25>最少6-10个字符</TD></TR>
        <TR>
          <TD width="21%" height=25>*确认：</TD>
          <TD width="31%" height=25><INPUT class=input type=password 
            maxLength=10 size=11 name=pswc></TD>
          <TD width="48%" height=25>再输入一次密码确认</TD></TR>
<TR>
          <TD width="21%" height=25>*第二密码：</TD>
          <TD width="31%" height=25><INPUT class=input type=password 
            maxLength=10 size=11 name=psw2></TD>
          <TD width="48%" height=25>用来找回丢失的密码</TD></TR>
<TR>
  
              <td height="25" width="71" align="center">国家<font color="#FF0080"> 
              *</font></td>       
              <td height="25" width="276"> &nbsp;        
                    <select name="country"> 
                  <option value="无" selected>自由者</option>                
                  <option value="秦国">秦国</option>
                  <option value="楚国">楚国</option>
                  <option value="赵国">赵国</option>
                  <option value="魏国">魏国</option>
                  <option value="齐国">齐国</option>
                  <option value="韩国">韩国</option>
                  <option value="燕国">燕国</option>
                </select></td>
       </tr>

                 <TD width="21%" height=25>*QQ：</TD>
          <TD width="31%" height=25><INPUT class=input maxLength=11 size=11 name="oicq"></TD>
          <TD width="48%" height=25>填写QQ好与朋友联系5位以上</TD></TR>
        <TR>
          <TD width="21%" height=25>*电子信箱：</TD>
          <TD width="31%" height=25><INPUT class=input maxLength=25 size=11 
            name=e_mail value=""></TD>
          <TD width="48%" height=25>正确填写真实信箱</TD></TR>        <TR>
          <TD width="21%" height=25>介绍人：</TD>
          <TD width="31%" height=25><INPUT class=input maxLength=5 size=11 name=jsr value=""></TD>
          <TD width="48%" height=25>介绍人，没有请留空</TD></TR>
        <TR>
          <TD width="21%" height=25>*认证：</TD>
          <TD width="31%" height=25><input type=hidden name=regjm1 value="<%=regjm%>"><INPUT class=input maxLength=4 size=11 
            name=regjm2 value='<%=left(regjm,4)%>' onKeypress="if (event.keyCode < 45 || event.keyCode > 57) {event.returnValue = false;}"></TD>
          <TD width="48%" height=25>认证请输入：<FONT color=red><%=regjm%></FONT></TD></TR>
        <TR>
          <TD align=middle width="100%" colSpan=3 height=29><INPUT style="BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid; BACKGROUND-COLOR: #e8e8d8" type=submit value="申请" name="submit"> 
<INPUT style="BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid; BACKGROUND-COLOR: #e8e8d8" onclick=window.close() type=button value="关闭" name=button></form></TD></TR>
        <TR>
          <TD width="100%" colSpan=3 height=21><FONT 
            color=#ff0000>注意事项：请牢记你所填写的帐号和密码，请填上介绍</FONT></TD></TR>
        <TR>
          <TD width="100%" colSpan=3 height=17><FONT 
            color=#ff0000>你来这个江湖的人的名字，以便对其适当奖励</FONT></TD></TR>
        <TR>
          <TD align=middle colSpan=3><FONT 
            color=#0000ff>&copy; 版权所有 2004-2008 </FONT><a href="http://www.happyjh.com/" target="_blank"><font color="#0000ff">快乐江湖网</font></a> </TD></TR>  
</table><!=></TD></TR>
<TR><TD width="3%" height=6><IMG src="pic/T3.gif" border=0></TD>
    <TD width="94%" background=pic/M3.gif height=6>　</TD>
    <TD width="12%" height=6><IMG src="pic/T4.gif" 
  border=0></TD></TR></TBODY></TABLE></BODY></HTML>