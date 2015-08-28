<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then
%>
<script language=vbscript>
MsgBox "对不起，你还没有登录剑侠！"
window.close()
</script>
<%
else
action=session("yx8_mhjh_username")
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="Select  * from 用户 where 姓名='"& action &"'"
set rs=conn.Execute(sql)
%>
<html>

<head>
<title>档案修改</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<STYLE type=text/css></STYLE>

<link rel="stylesheet" href="../style.css" type="text/css">
</head>

<body oncontextmenu=self.event.returnValue=false topmargin="0" leftmargin="0" bgcolor="#000000" >
<form method="POST" action="regmodi1.asp">
  <div align="center">
    <table width="580" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td background="../image/bg1.gif">
          <div align="center">
            <table border="0" width="475" cellspacing="0" cellpadding="1" align="center">
              <tr> 
                <td colspan="6" height="22"> 
                  <div align="center"><b><%=rs("姓名")%></b><font color="#924949"><b>大侠的背景</b></font></div>
                </td>
              </tr>
              <tr > 
                <td width="49" height="26">姓 名:</td>              
                <td height="26" width="154"><%=rs("姓名")%></td>              
                <td height="26" width="42">头衔：</td>              
                <td height="26" width="74">               
                  <input type="text"                         
name="touxian" size="8"                         
style="font-family: Tahoma; font-size: 12px"                         
maxlength="50" value="<%=rs("头衔")%>">              
                  　</td>              
                <td height="26" width="85"> 荣誉：</td>              
                <td height="26" width="59">               
                  <%                         
if rs("等级")="" then response.write("武林新手")                         
if rs("等级")<=3 then response.write("武林新手")                         
if rs("等级")<=6 and rs("等级")>3  then response.write("武林游侠")                         
if rs("等级")<=8 and rs("等级")>6  then response.write("武林斗侠")                         
if rs("等级")<=10and rs("等级")>8  then response.write("初级战将")                         
if rs("等级")<=11 and rs("等级")>10  then response.write("中级战将")                         
if rs("等级")<=14 and rs("等级")>11  then response.write("高级战将")                         
if rs("等级")<=17 and rs("等级")>14  then response.write("超级战将")                         
if rs("等级")>17 then response.write("终极战神")                         
%>              
                  　</td>              
              </tr>              
              <tr >               
                <td width="49" height="18"><font class="c">性 别:</font></td>              
                <td height="18" width="154">               
                  <%if rs("性别")="男"  then response.write("帅哥")                         
if rs("性别")="女"  then response.write("美女")                         
%>              
                </td>              
                <td height="18" width="42">状态：</td>              
                <td height="18" width="74"><%=rs("状态")%></td>              
                <td height="18" width="85">师傅/徒弟：</td>              
                <td height="18" width="59"><%=rs("师傅")%>/<%=rs("徒弟")%></td>              
              </tr>              
              <tr >               
                <td width="49"><font class="c">Email:</font></td>              
                <td nowrap width="154">               
                  <input type="text"                         
name="email" size="20"                         
style="font-family: Tahoma; font-size: 12px"                         
maxlength="50" value="<%=rs("电子邮箱")%>">              
                </td>              
                <td width="42">OICQ：</td>              
                <td width="74">               
                  <input type="text"                         
name="oicq" size="8"                         
style="font-family: Tahoma; font-size: 12px"                         
maxlength="9" value="<%=rs("oicq")%>">              
                </td>              
                <td width="85">结拜兄弟 ：</td>              
                <td width="59"><%=rs("结拜兄弟")%></td>              
              </tr>              
            </table>              
            <table border="0"                         
width="471"  cellspacing="0" cellpadding="1" align="center">              
              <tr >               
                <td colspan="5" height="20">               
                  <div align="center">               
                    <p>&nbsp;</p>              
                    <p><font class="p9" size="2" ><b><font color="#924949">江 湖               
                      档 案</font></b></font> </p>              
                  </div>              
                </td>              
              </tr>              
              <tr >              
                <td width="5" height="25">&nbsp;</td>              
                <td width="54" height="25">现 金：</td>              
                <td width="199" height="25"><%=rs("银两")%> 两</td>              
                <td width="65" height="25">积 分：</td>              
                <td height="25" width="173"><%=rs("积分")%></td>              
              </tr>              
              <tr >              
                <td width="5" height="20">&nbsp;</td>              
                <td width="54" height="20">存 款：</td>              
                <td width="199" height="24"><%=rs("存款")%> 两</td>              
                <td width="65" height="24">道 德：</td>              
                <td height="24" width="173"><%=rs("道德")%></td>              
              </tr>              
              <tr >              
                <td width="5" height="20">&nbsp;</td>              
                <td width="54" height="20">武 功：</td>              
                <td width="199" height="24"><%=rs("防御")%></td>              
                <td width="65" height="24">攻 击：</td>              
                <td height="24" width="173"><%=rs("攻击")%></td>              
              </tr>              
              <tr >              
                <td width="5" height="20">&nbsp;</td>              
                <td width="54" height="20">内 力：</td>              
                <td width="199"><%=rs("内力")%></td>              
                <td width="65">防 御：</td>              
                <td width="173"><%=rs("防御")%></td>              
              </tr>              
              <tr >              
                <td width="5" height="20">&nbsp;</td>              
                <td width="54" height="20">精 力：</td>              
                <td width="199"><%=rs("精力")%></td>              
                <td width="65">门 派：</td>              
                <td width="173"><%=rs("门派")%> </td>              
              </tr>              
              <tr >              
                <td width="5" height="20">&nbsp;</td>              
                <td width="54" height="20">体 力：</td>              
                <td width="199"><%=rs("体力")%></td>              
                <td width="65">身 份：</td>              
                <td width="173"><%=rs("身份")%></td>              
              </tr>              
              <tr >              
                <td width="5" height="20">&nbsp;</td>              
                <td width="54" height="20">配 偶：</td>              
                <td width="199"><%=rs("配偶")%></td>              
                <td width="65">杀 人：</td>              
                <td width="173"><%=rs("总杀人")%></td>              
              </tr>              
              <tr >              
                <td width="5" height="20">&nbsp;</td>              
                <td width="54" height="20">等 级：</td>              
                <td width="199"> <%=rs("等级")%>级 </td>              
                <td width="65">被 杀：</td>              
                <td width="173"><%=rs("被杀")%>　</td>              
              </tr>              
              <tr >               
                <td colspan="5" height="199">               
                  <table width="480" border="0" cellspacing="0" cellpadding="0">              
                    <tr>               
                      <td height="196" width="270">               
                        <p align="center">你当前的头像是： <font color="#000000"><img src="<%=rs("头像")%>">               
                          </font></p>              
                        <p align="center">下面请选择你喜欢的头像</p>              
                        <div align="center"><font color="#000000" size="2"> </font>               
                          <font color="#000000"><img id=face src="../face/001.gif" alt=个人形象代表 width="36" height="31">&nbsp;               
                          <select name=touxiang size=1 onChange="document.images['face'].src=options[selectedIndex].value;">              
                            <option selected value="../face/001.gif">形象1 
                            <option value="../face/002.gif">形象2 
                            <option value="../face/003.gif">形象3 
                            <option value="../face/004.gif">形象4 
                            <option value="../face/005.gif">形象5 
                            <option value="../face/006.gif">形象6 
                            <option value="../face/007.gif">形象7 
                            <option value="../face/008.gif">形象8 
                            <option value="../face/009.gif">形象9 
                            <option value="../face/010.gif">形象10 
                            <option value="../face/012.gif">形象11 
                            <option value="../face/013.gif">形象12 
                            <option value="../face/014.gif">形象13 
                            <option value="../face/015.gif">形象14 
                            <option value="../face/016.gif">形象15 
                            <option value="../face/017.gif">形象16 
                            <option value="../face/018.gif">形象17 
                            <option value="../face/019.gif">形象18 
                            <option value="../face/020.gif">形象19 
                            <option value="../face/021.gif">形象20 
                            <option value="../face/022.gif">形象21 
                            <option value="../face/023.gif">形象22 
                            <option value="../face/024.gif">形象23 
                            <option value="../face/025.gif">形象24 
                            <option value="../face/026.gif">形象25 
                            <option value="../face/027.gif">形象26 
                            <option value="../face/028.gif">形象27 
                            <option value="../face/029.gif">形象28 
                            <option value="../face/030.gif">形象29 
                            <option value="../face/031.gif">形象30 
                            <option value="../face/032.gif">形象31 
                            <option value="../face/033.gif">形象32 
                            <option value="../face/034.gif">形象33 
                            <option value="../face/035.gif">形象34 
                            <option value="../face/036.gif">形象35 
                            <option value="../face/037.gif">形象36 
                            <option value="../face/038.gif">形象37 
                            <option value="../face/039.gif">形象38 
                            <option value="../face/040.gif">形象39 
                            <option value="../face/041.gif">形象40 
                            <option value="../face/042.gif">形象41 
                            <option value="../face/043.gif">形象42 
                            <option value="../face/044.gif">形象43 
                            <option value="../face/045.gif">形象44 
                            <option value="../face/046.gif">形象45 
                            <option value="../face/047.gif">形象46 
                            <option value="../face/048.gif">形象47 
                            <option value="../face/049.gif">形象48 
                            <option value="../face/050.gif">形象49 
                            <option value="../face/051.gif">形象50 
                            <option value="../face/052.gif">形象51 
                            <option value="../face/053.gif">形象52 
                            <option value="../face/054.gif">形象53 
                            <option value="../face/055.gif">形象54 
                            <option value="../face/056.gif">形象55 
                            <option value="../face/057.gif">形象56 
                            <option value="../face/058.gif">形象57 
                            <option value="../face/059.gif">形象58 
                            <option value="../face/060.gif">形象59 
                            <option value="../face/061.gif">形象60 
                            <option value="../face/062.gif">形象61 
                            <option value="../face/063.gif">形象62 
                            <option value="../face/064.gif">形象63 
                            <option value="../face/065.gif">形象64 
                            <option value="../face/066.gif">形象65 
                            <option value="../face/067.gif">形象66 
                            <option value="../face/068.gif">形象67 
                            <option value="../face/069.gif">形象68 
                            <option value="../face/070.gif">形象69 
                            <option value="../face/071.gif">形象70 
                            <option value="../face/072.gif">形象71 
                            <option value="../face/073.gif">形象72 
                            <option value="../face/074.gif">形象72 
                            <option value="../face/075.gif">形象74 
                            <option value="../face/076.gif">形象75 
                            <option value="../face/077.gif">形象76 
                            <option value="../face/078.gif">形象77 
                            <option value="../face/079.gif">形象78 
                            <option value="../face/080.gif">形象79 
                            <option value="../face/081.gif">形象80 
                            <option value="../face/082.gif">形象81 
                            <option value="../face/083.gif">形象82 
                            <option value="../face/084.gif">形象83 
                            <option value="../face/085.gif">形象84 
                            <option value="../face/086.gif">形象85 
                            <option value="../face/087.gif">形象86 
                            <option value="../face/088.gif">形象87 
                            <option value="../face/089.gif">形象88 
                            <option value="../face/090.gif">形象89 
                            <option value="../face/091.gif">形象90 
                            <option value="../face/092.gif">形象91 
                            <option value="../face/093.gif">形象92 
                            <option value="../face/094.gif">形象93 
                            <option value="../face/095.gif">形象94 
                            <option value="../face/096.gif">形象95 
                            <option value="../face/097.gif">形象96 
                            <option value="../face/098.gif">形象97 
                            <option value="../face/099.gif">形象98 
                            <option value="../face/100.gif">形象99 
                          </select>
                          <br>
                          </font> </div>
                        <div align="center">　<br>
                          <br>
                        </div>
                        <div align="center"> 
                          <p>注意：修改完成后请重新登陆江湖！ </p>
                        </div>
                      </td>
                      <td height="196" width="204" align="left" valign="top"> 
                        <div align="center">个人签名，写上想与朋友说的话！<span class="txt"> 
                          <textarea name="qianming" cols="30" style="font-family: Tahoma; font-size: 12px" rows="8"><%=rs("签名档")%></textarea>              
                          </span></div>              
                      </td>              
                    </tr>              
                  </table>              
                  <div align="right">               
                    <p>&nbsp;</p>              
                    <p align="center">               
                      <input type="submit" value="我要修改" name="B1" tyle="font-family: Tahoma; font-size: 12px">              
                    </p>              
                    <p align="center">&nbsp; </p>              
                  </div>              
                </td>              
              </tr>              
            </table>              
          </div>              
        </td>              
      </tr>              
    </table>              
  </div>              
  <div align="left"></div>                         
  <div align="center"> </div>                         
</form>                         
<div align="center">                         
<br>                         
</div>                         
</body>                         
                         
</html>                         
<%           
end if                          
rs.close                         
set rs=nothing         
conn.close                         
set conn=nothing                                           
%>                         
