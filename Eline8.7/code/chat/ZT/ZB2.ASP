<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../config.asp"-->

<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.solbutton {background-image: url(../images/bgbutton.gif); background-repeat:repeat-x; background-color: #FFCF4A; color=black;width=60;height=23; border=#FF0000 ;border-bottom: 1px solid #996633; border-left: 1px solid #996633; border-right: 1px solid #996633; border-top: 1px solid #996633;cursor=hand;}
TD {
	FONT-SIZE: 9pt
}
-->
</style>
<title>♀wWw.happyjh.com♀</title>
<SCRIPT language=javascript>
function showwp(ename,cname,tx,tj,lb,liliang,fangyu,sudu1,sudu,price){
	zl.innerHTML='<input type=hidden name=wpname value='+cname+'>'+cname+'[<font color="#FFFF00">'+lb+'</font>]<br>攻击：'+liliang+'<br>防御：'+fangyu+'<br>耐久：<font color=red>'+sudu1+'</font>/<font color=blue>'+sudu+'</font><br>特性:'+tx+'<br>'
}
function yesno(name) {
if(confirm("您现在将要卸下的是 ["+name+"]  \n您现在可以点确定卸下。")) {document.form1.submit();}
}
</SCRIPT>
</head>
<%
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
function gjl(lb)
if isnull(rs(lb)) then
	gjl=0
	exit function
end if
data1=split(rs(lb),"|")
rs.close
rs.open "select * FROM b WHERE a='" & data1(0) &"'",conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	gjl=0
	exit function
end if
gjl=int(rs("f")*(clng(data1(1))/rs(j)))
'myzb="<img src='../../hcjs/JHJS/images/"&rs("i")&"' alt='"&rs("c")&"'>["&rs("a")&"] 攻击:"&rs("f")&" 防御:"&rs("g")&" 耐久:<font color=red>"&data1(1)&"</font>/<font color=blue>"&rs("j")&"</font> 特性:"&rs("k")
rs.close
end function
function myzb(lb)
if isnull(rs(lb)) or trim(rs(lb))="" then
	rs.close
	myzb="无"
	exit function
end if
data1=split(rs(lb),"|")
rs.close
rs.open "select * FROM b WHERE a='" & data1(0) &"'",conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	myzb="无"
	exit function
end if
myzb="<A href='#' onclick="&chr(34)&"showwp('"&rs("i")&"','"&rs("a")&"','"&rs("k")&"','"&tj&"','"&rs("b")&"',"&rs("f")&","&rs("g")&","&data1(1)&","&rs("j")&","&rs("h")&");"&chr(34)&" title='"&rs("c")&"'><img src='../../hcjs/jhjs/images/"&rs("i")&"'width='39' alt='"&rs("c")&"'></A>"
rs.close
end function
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,3,3
tlj=rs("体力加")
nlj=rs("内力加")
%>
<body topmargin="0" leftmargin="0" text="#FFFFFF" style="font-size: 9pt" bgcolor="#CCCCCC" oncontextmenu=self.event.returnValue=false>
<table height="428" border="0" cellpadding="0" cellspacing="0">
  <FORM action=xxia.asp method=post name=form1>
    <tr> 
      <td><img src="img/shim.gif" width="11" height="1" border="0"></td>
      <td><img src="img/shim.gif" width="30" height="1" border="0"></td>
      <td><img src="img/shim.gif" width="6" height="1" border="0"></td>
      <td><img src="img/shim.gif" width="31" height="1" border="0"></td>
      <td><img src="img/shim.gif" width="7" height="1" border="0"></td>
      <td><img src="img/shim.gif" width="6" height="1" border="0"></td>
      <td><img src="img/shim.gif" width="1" height="1" border="0"></td>
      <td><img src="img/shim.gif" width="3" height="1" border="0"></td>
      <td><img src="img/shim.gif" width="10" height="1" border="0"></td>
      <td><img src="img/shim.gif" width="34" height="1" border="0"></td>
      <td><img src="img/shim.gif" width="10" height="1" border="0"></td>
      <td><img src="img/shim.gif" width="5" height="1" border="0"></td>
      <td><img src="img/shim.gif" width="7" height="1" border="0"></td>
      <td><img src="img/shim.gif" width="5" height="1" border="0"></td>
      <td><img src="img/shim.gif" width="38" height="1" border="0"></td>
      <td><img src="img/shim.gif" width="6" height="1" border="0"></td>
      <td><img src="img/shim.gif" width="23" height="1" border="0"></td>
      <td><img src="img/shim.gif" width="10" height="1" border="0"></td>
      <td><img src="img/shim.gif" width="10" height="1" border="0"></td>
      <td><img src="img/shim.gif" width="1" height="1" border="0"></td>
    </tr>
    <tr valign="top"> 
      <td colspan="20"><img src="img/shim.gif" ></td>
    </tr>
    <tr valign="top"> 
      <td colspan="19" rowspan="13" bgcolor="#2B1515"> 
        <table border="0" cellpadding="0" cellspacing="0" width="255">
          <!-- fwtable fwsrc="未命名" fwbase="1.jpg" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
          <tr> 
            <td><img src="spacer.gif" width="35" height="1" border="0"></td>
            <td><img src="spacer.gif" width="6" height="1" border="0"></td>
            <td><img src="spacer.gif" width="39" height="1" border="0"></td>
            <td><img src="spacer.gif" width="6" height="1" border="0"></td>
            <td><img src="spacer.gif" width="9" height="1" border="0"></td>
            <td><img src="spacer.gif" width="36" height="1" border="0"></td>
            <td><img src="spacer.gif" width="9" height="1" border="0"></td>
            <td><img src="spacer.gif" width="6" height="1" border="0"></td>
            <td><img src="spacer.gif" width="6" height="1" border="0"></td>
            <td><img src="spacer.gif" width="10" height="1" border="0"></td>
            <td><img src="spacer.gif" width="29" height="1" border="0"></td>
            <td><img src="spacer.gif" width="6" height="1" border="0"></td>
            <td><img src="spacer.gif" width="10" height="1" border="0"></td>
            <td><img src="spacer.gif" width="48" height="1" border="0"></td>
            <td><img src="spacer.gif" width="1" height="1" border="0"></td>
          </tr>
          <tr> 
            <td colspan="3"><img name="n1_r1_c1" src="1_r1_c1.jpg" width="80" height="29" border="0"></td>
            <td rowspan="2" colspan="5"><img name="n1_r1_c4" src="1_r1_c4.jpg" width="66" height="58" border="0"></td>
            <td rowspan="2" colspan="3"><img name="n1_r1_c9" src="1_r1_c9.jpg" width="45" height="58" border="0"></td>
            <td rowspan="2" colspan="3"><img name="n1_r1_c12" src="1_r1_c12.jpg" width="64" height="58" border="0"></td>
            <td><img src="spacer.gif" width="1" height="29" border="0"></td>
          </tr>
          <tr> 
            <td rowspan="2"><img name="n1_r2_c1" src="1_r2_c1.jpg" width="35" height="45" border="0"></td>
            <td colspan="2" rowspan="2" background="1_r2_c2.jpg">
              <%
            rs.close
            rs.open "SELECT z1 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,3,3
            %>
              <%=myzb("z1")%></td>
            <td><img src="spacer.gif" width="1" height="29" border="0"></td>
          </tr>
          <tr> 
            <td rowspan="3" colspan="5"><img name="n1_r3_c4" src="1_r3_c4.jpg" width="66" height="63" border="0"></td>
            <td colspan="3" rowspan="2" background="1_r2_c2.jpg">
              <%rs.open "SELECT z2 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,3,3 %>
              <%=myzb("z2")%></td>
            <td rowspan="5" colspan="3"><img name="n1_r3_c12" src="1_r3_c12.jpg" width="64" height="72" border="0"></td>
            <td><img src="spacer.gif" width="1" height="16" border="0"></td>
          </tr>
          <tr> 
            <td rowspan="4"><img name="n1_r4_c1" src="1_r4_c1.jpg" width="35" height="56" border="0"></td>
            <td rowspan="4" colspan="2"><img name="n1_r4_c2" src="1_r4_c2.jpg" width="45" height="56" border="0"></td>
            <td><img src="spacer.gif" width="1" height="29" border="0"></td>
          </tr>
          <tr> 
            <td rowspan="3" colspan="3"><img name="n1_r5_c9" src="1_r5_c9.jpg" width="45" height="27" border="0"></td>
            <td><img src="spacer.gif" width="1" height="18" border="0"></td>
          </tr>
          <tr> 
            <td rowspan="2" colspan="2"><img name="n1_r6_c4" src="1_r6_c4.jpg" width="15" height="9" border="0"></td>
            <td colspan="2" rowspan="3" background="1_r2_c2.jpg">
              <%rs.open "SELECT z3 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,3,3 %>
              <%=myzb("z3")%></td>
            <td rowspan="6"><img name="n1_r6_c8" src="1_r6_c8.jpg" width="6" height="147" border="0"></td>
            <td><img src="spacer.gif" width="1" height="1" border="0"></td>
          </tr>
          <tr> 
            <td><img src="spacer.gif" width="1" height="8" border="0"></td>
          </tr>
          <tr> 
            <td rowspan="2" colspan="2"><img name="n1_r8_c1" src="1_r8_c1.jpg" width="41" height="45" border="0"></td>
            <td colspan="2" rowspan="2" background="1_r2_c2.jpg">
              <%rs.open "SELECT z5 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,3,3 %>
              <%=myzb("z5")%></td>
            <td rowspan="4"><img name="n1_r8_c5" src="1_r8_c5.jpg" width="9" height="138" border="0"></td>
            <td rowspan="4"><img name="n1_r8_c9" src="1_r8_c9.jpg" width="6" height="138" border="0"></td>
            <td colspan="3" rowspan="2" background="1_r2_c2.jpg">
              <%rs.open "SELECT z4 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,3,3 %>
              <%=myzb("z4")%></td>
            <td rowspan="2" colspan="2"><img name="n1_r8_c13" src="1_r8_c13.jpg" width="58" height="45" border="0"></td>
            <td><img src="spacer.gif" width="1" height="36" border="0"></td>
          </tr>
          <tr> 
            <td rowspan="3" colspan="2"><img name="n1_r9_c6" src="1_r9_c6.jpg" width="45" height="102" border="0"></td>
            <td><img src="spacer.gif" width="1" height="9" border="0"></td>
          </tr>
          <tr> 
            <td rowspan="2" colspan="2"><img name="n1_r10_c1" src="1_r10_c1.jpg" width="41" height="93" border="0"></td>
            <td rowspan="2" colspan="2"><img name="n1_r10_c3" src="1_r10_c3.jpg" width="45" height="93" border="0"></td>
            <td rowspan="2" colspan="3"><img name="n1_r10_c10" src="1_r10_c10.jpg" width="45" height="93" border="0"></td>
            <td colspan="2"><img name="n1_r10_c13" src="1_r10_c13.jpg" width="58" height="46" border="0"></td>
            <td><img src="spacer.gif" width="1" height="46" border="0"></td>
          </tr>
          <tr> 
            <td colspan="2"><img name="n1_r11_c13" src="1_r11_c13.jpg" width="58" height="47" border="0"></td>
            <td><img src="spacer.gif" width="1" height="47" border="0"></td>
          </tr>
          <tr> 
            <td rowspan="2" colspan="2"><img name="n1_r12_c1" src="1_r12_c1.jpg" width="41" height="81" border="0"></td>
            <td rowspan="2" colspan="2"><img name="n1_r12_c3" src="1_r12_c3.jpg" width="45" height="81" border="0"></td>
            <td rowspan="2" colspan="2"><img name="n1_r12_c5" src="1_r12_c5.jpg" width="45" height="81" border="0"></td>
            <td rowspan="2" colspan="4"><img name="n1_r12_c7" src="1_r12_c7.jpg" width="31" height="81" border="0"></td>
            <td colspan="3" background="1_r2_c2.jpg">
              <%rs.open "SELECT z6 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,3,3 %>
              <%=myzb("z6")%></td>
            <td><img name="n1_r12_c14" src="1_r12_c14.jpg" width="48" height="45" border="0"></td>
            <td><img src="spacer.gif" width="1" height="45" border="0"></td>
          </tr>
          <tr> 
            <td colspan="3"><img name="n1_r13_c11" src="1_r13_c11.jpg" width="45" height="36" border="0"></td>
            <td><img name="n1_r13_c14" src="1_r13_c14.jpg" width="48" height="36" border="0"></td>
            <td><img src="spacer.gif" width="1" height="36" border="0"></td>
        </table>
      <td><img src="img/shim.gif" width="1" height="31" border="0"></td>
    </tr>
    <tr valign="top"> 
      <td><img src="img/shim.gif" width="1" height="5" border="0"></td>
    </tr>
    <tr valign="top"> 
      <td><img src="img/shim.gif" width="1" height="38" border="0"></td>
    </tr>
    <tr valign="top"> 
      <td><img src="img/shim.gif" width="1" height="5" border="0"></td>
    </tr>
    <tr valign="top"> 
      <td><img src="img/shim.gif" width="1" height="11" border="0"></td>
    </tr>
    <tr valign="top"> 
      <td><img src="img/shim.gif" width="1" height="16" border="0"></td>
    </tr>
    <tr valign="top"> 
      <td><img src="img/shim.gif" width="1" height="27" border="0"></td>
    </tr>
    <tr valign="top"> 
      <td><img src="img/shim.gif" width="1" height="17" border="0"></td>
    </tr>
    <tr valign="top"> 
      <td><img src="img/shim.gif" width="1" height="44" border="0"></td>
    </tr>
    <tr valign="top"> 
      <td><img src="img/shim.gif" width="1" height="39" border="0"></td>
    </tr>
    <tr valign="top"> 
      <td><img src="img/shim.gif" width="1" height="24" border="0"></td>
    </tr>
    <tr valign="top"> 
      <td><img src="img/shim.gif" width="1" height="43" border="0"></td>
    </tr>
    <tr valign="top"> 
      <td><img src="img/shim.gif" width="1" height="60" border="0"></td>
    </tr>
    <tr valign="top"> 
      <td rowspan="4"><img name="zb_r15_c01" src="img/zb_r15_c01.jpg" width="11" height="84" border="0"></td>
      <td id=zl vAlign=top colspan="11" background="img/zb_r15_c02.jpg" rowspan="3"><input type=hidden name=wpname value=''>
        　</td>
      <td  colspan="6" background="img/zb_r15_c02.jpg">　</td>
      <td rowspan="4"><img name="zb_r15_c19" src="img/zb_r15_c19.jpg" width="10" height="84" border="0"></td>
      <td rowspan="3"><img src="img/shim.gif" width="1" height="72" border="0"></td>
    </tr>
    <tr valign="top"> 
      <td  colspan="6" background="img/zb_r15_c02.jpg"><input type="submit"  value="卸  下" class=solbutton onclick="if(document.form1.wpname.value!=''){yesno(document.form1.wpname.value);return false;}else{return false;}"> 
      </td>
    </tr>
    <tr valign="top"> 
      <td  colspan="6" background="img/zb_r15_c02.jpg"> <input onClick="javascript:window.document.location.href='daojun.asp'" class=solbutton type=button value="装备物品" name="button"> 
      </td>
    </tr>
    <tr valign="top"> 
      <td colspan="17"><img name="zb_r16_c02" src="img/zb_r16_c02.jpg" width="232" height="12" border="0"></td>
      <td><img src="img/shim.gif" width="1" height="12" border="0"></td>
    </tr>
  </FORM>
</table>
</body>
</html>