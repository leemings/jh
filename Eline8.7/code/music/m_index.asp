<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
//if sjjh_name="" then Response.Redirect "../error.asp?id=440"
dim membername
membername=session("sjjh_name")
%>
<!--#include file="conn.asp"-->
<!--#include file="Top.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="images/home.css" type="text/css">
<title>一线视听 → wWw.51Eline.com</title>
<style type="text/css">
BODY {CURSOR: url('images/mouse.ani');
scrollbar-face-color:#ededed; 
scrollbar-shadow-color:#cccccc; 
scrollbar-highlight-color:#efefef;
scrollbar-3dlight-color:#ededed;
scrollbar-darkshadow-color:#ededed;
scrollbar-track-color:#ffffff;
scrollbar-arrow-color:#0993f4;
}
</style>
</head>

<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="766">
    <tr>
      <td width="100%" bgcolor="#F79618">
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="27">
            <tr bgcolor="#66CCFF"> 
              <td width="5%" height="27"> 
                <p align="right">
                          <font style="font-size: 9pt">
              <img border="0" src="images/oo.gif"></font></td>
              <td width="7%" height="27"> 
<form name="search" method="post" action="search.asp">
                          <p align="left">
                          <font style="font-size: 9pt">&nbsp;搜索：</font></p>
                </td>
                <td width="13%" height="27"> <span style="font-size: 9pt"> 
                  <input name="keyword" size="14" style="border-style: solid; border-width: 1; background-color: #FFFFFF; font-size:9pt"></span></td>
                <td width="6%" height="27"> 
                  <p align="center"> 
                            <span style="font-size: 9pt">类别</span><font style="font-size: 9pt">：</font></td>
                <td width="5%" height="27"><span style="font-size: 9pt"> 
                  <select name=stype style="color:000088">
<option value="Music" selected>歌曲名称</option>
<option value="Special">专辑名称</option>
<option value="Singer">歌手姓名</option> 
</select></span></td>
                <td width="1%" height="27">　</td>
                <td width="7%" height="27"> <font style="font-size: 9pt"> 
                  <input type=image src=images/search.gif name="I3"></font></td>
                <td width="57%" height="27"> <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber43"> 
                  <tr></form>
                  
              <td width="38%">
<td width="61%"> 
                <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber48">
                    <tr>
                      <td width="32%" align="center"></td>
                      <td width="36%" align="center">
                      <img border="0" src="images/forum_readme.gif"><A href="#" 
      onmousedown="window.external.addFavorite('http://www.51Eline.com','≮一线网络≯-wWw.51Eline.COM')" 
      target="_blank" title="一线网络">加入收藏夹</A></td>
                      <td width="32%" align="center">
                      <img border="0" src="images/forum_readme.gif"> <A href="#" 
      onmousedown="this.style.behavior='url(#default#homepage)';this.setHomePage('http://www.51Eline.com');" 
      style="BEHAVIOR: url(#default#homepage)" url(#default#homepage)>设置为首页</A></td>
                    </tr>
                  </table>
                  </td>
                  
              <td width="1%">　</td>
                </tr>
              </table>
              </td>
            </tr>
      </table>
      </td>
    </tr>
  </table>
  </center>
</div>

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="766" bgcolor="#FFFFFF" height="390">
    <tr>
      <td width="488" height="390" valign="top">
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="77%">
        <tr>
          <td width="100%">
          <table border="0" cellpadding="0" cellspacing="5" style="border-collapse: collapse" bordercolor="#111111" width="74%" height="324">
            <tr>
              <td width="100%" height="200" bgcolor="#FFFFFF">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="85%">
                <tr>
                  <td width="1%"><img border="0" src="images/io4.gif"></td>
                  <td width="99%" valign="top">
                  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#006194" width="100%">
                    <tr>
                  <td width="100%" height="1" bgcolor="#FFFCF0">
                  <table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="454" id="AutoNumber9" height="1">
                    <tr>
                      <td width="96" height="1"><%
			set rs=server.createobject("adodb.recordset")
			i=0 
			
			sql="SELECT * FROM Special order by Specialid desc"
			rs.Open sql,conn,1,1
			%>
			<%
			do while not rs.eof
				i=i+1
			%>
              <tr> 
                <td width="96" height="1"> 
                  <table width="106" border="0" cellspacing="1" cellpadding="0" height="95">
                    <tr> 
                      <td width="104"> 
                        <div align="center">
                          <center> 
                        <table width="80" border="0" cellspacing="0" cellpadding="0" height="80" style="border-collapse: collapse" bordercolor="#111111">
                          <tr>
                            <td><a href="MusicList.asp?AlbumID=<%=rs("SpecialID")%>">
                            <img height=80 src="<%=rs("pic")%>" width=80 border=1 style="border: 1px solid #000000"></a><img border="0" src="images/showdown.gif"><font color="006194"><br>
                            <img border="0" src="images/showdown2.gif" width="87" height="5"></font><br><%=rs("NClass")%></td>
                          </tr>
                        </table>
                          </center>
                        </div>
                      </td>
                    </tr>
                  </table>
                </td>
                <td class=a width="122" height="1"><font color="006194"><img border="0" src="images/img_arrow03.gif"><a href="MusicList.asp?AlbumID=<%=rs("SpecialID")%>"><b><%=rs("name")%></b> 
                  </a><br>
                  歌手：<%=rs("NClass")%><br>
                  语言：<%=rs("yuyan")%><br>
                  日期：<%=rs("times")%><br>
                  唱片公司：<%=rs("gongsi")%><a href="MusicList.asp?AlbumID=<%=rs("SpecialID")%>">
                  <%rs.movenext%>
                  </a></font></td>
                <td width="97" height="1"> 
                  <table width="95" border="0" cellspacing="1" cellpadding="0" height="95">
                    <tr>
                            <td><a href="MusicList.asp?AlbumID=<%=rs("SpecialID")%>">
                            <img height=80 src="<%=rs("pic")%>" width=80 border=1 style="border: 1px solid #000000"></a><img border="0" src="images/showdown.gif"><font color="006194"><br>
                            <img border="0" src="images/showdown2.gif" width="87" height="5"></font><br><%=rs("NClass")%></td>
                          </tr>
                  </table>
                </td>
                <td height="1" valign="middle" width="139"><font color="006194"><img border="0" src="images/img_arrow03.gif"><a href="MusicList.asp?AlbumID=<%=rs("SpecialID")%>"><b><%=rs("name")%></b> 
                  </a><br>
                  歌手：<%=rs("NClass")%><br>
                  语言：<%=rs("yuyan")%><br>
                  日期：<%=rs("times")%><br>
                  唱片公司：<%=rs("gongsi")%><a href="MusicList.asp?AlbumID=<%=rs("SpecialID")%>">
                  <%rs.movenext%>
                  </a></font></td>
              </tr>
			  <%
				if i>=2 then exit do
				rs.movenext
				loop
			  %>
                  </table>
                  <%
set rs=server.createobject("adodb.recordset")
sql="SELECT top 1 * FROM Special where IsGood=1 order by Specialid desc"
rs.Open sql,conn,1,1
%>

                  </td>
                    </tr>
                  </table>
                  </td>
                </tr>
              </table>
              </td>
            </tr>
            <tr>
              <td width="100%" height="16" bgcolor="#FFFFFF">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="1%" valign="top">
                  <p align="right"><img border="0" src="images/io5.gif"></td>
                  <td width="99%" valign="top">
                  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#004984" width="100%">
                    <tr>
                  <td width="100%" height="70" valign="top" bgcolor="#FFFCF0">
                  <table border="0" cellpadding="8" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="99%" id="AutoNumber13" height="1">
                    <tr>
                      <td width="24%" height="1" valign="top" bgcolor="#FFFFFF">
                      <table border="0" cellpadding="6" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="81%" id="AutoNumber14">
                        <tr>
                          <td width="100%"><a href="MusicList.asp?AlbumID=<%=rs("SpecialID")%>">
                          <img height=80 src="<%=rs("pic")%>" width=80 border=0 style="border: 1px solid #000000"></a></td>
                        </tr>
                        <tr>
                          <td width="100%">
                          <p align="center"><b><%=rs("NClass")%></b></td>
                        </tr>
                      </table>
                      </td>
                      <td width="76%" height="1" valign="top" rowspan="2" bgcolor="#FFFFFF">
                      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber15" height="37">
                        <tr>
                          <td width="11%" height="37" rowspan="2"><font color="006194"><img border="0" src="images/img_arrow01.gif"></td>
                          <td width="89%" height="22">歌手姓名：<%=rs("NClass")%></td>
                        </tr>
                        <tr>
                          <td width="89%" height="15">专集名称：<a href="MusicList.asp?AlbumID=<%=rs("SpecialID")%>"><%=rs("name")%></a></td>
                        </tr>
                        <tr>
                          <td width="100%" height="10" colspan="2"><%
if len(rs("intro"))>220 then
response.write (left(rs("intro"),220))
else
response.write (rs("intro"))
end if
%>......</td>
                        </tr>
                        <tr>
                          <td width="100%" height="6" colspan="2"></td>
                        </tr>
                        <tr>
                          <td width="100%" height="7" colspan="2">
                          <table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                            <tr>
                              <td width="100%"><%
m=0
set rs=conn.execute("SELECT top 1 ID,Specialid,NClassID FROM Musiclist order by ID desc") 
if not Rs.eof then
do while not rs.eof
%> 本站目前共有: 歌手<font color="red">1900</font>位&nbsp;&nbsp;专辑<font color="red"><%=rs("Specialid")%></font>张&nbsp;&nbsp;歌曲<font color="red"><%=rs("ID")%></font>首    
<%
if m>=1 then exit do
rs.movenext 
loop
end if
rs.close
%></td>
                            </tr>
                          </table>
                          </td>
                        </tr>
                      </table>
                      </td>
                    </tr>
                    <tr>
                      <td width="24%" height="1" valign="top" bgcolor="#FFFFFF">
                      <p align="center"><img border="0" src="images/1arrow.gif"></td>
                      </tr>
                  </table>
                  </td>
                    </tr>
                  </table>
                  </td>
                </tr>
              </table>
              </td>
            </tr>
            <tr>
              <td width="100%" height="67" bgcolor="#FFFFFF">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
          <td width="100%" valign="bottom">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="58%">
              <marquee>本站音乐全为WMA格式如不能正常播放,请下载升级播放程序，方可收听</marquee></td>
              <td width="42%">
              <p align="right"><img border="0" src="images/io8.gif"></td>
            </tr>
          </table>
                  </td>
                </tr>
                <tr>
          <td width="100%">
          <table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#004984" width="100%" id="AutoNumber22" height="1">
                <tr bgcolor="#FCCF95" align="center"> 
                  <td width="481" height="16" bgcolor="#004984" colspan="6">
                  <table border="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber29">
                    <tr>
                      <td width="2%" height="1">&nbsp; </td>
                      <td width="43%" height="1" align="center">
                      <font color="#FFFCF0">推荐歌曲</font></td>
                      <td width="19%" height="1" align="center">
                      <font color="#FFFCF0">演唱歌手</font></td>
                      <td width="13%" height="1" align="center">
                      <font color="#FFFCF0">WMA试听</font></td>
                      <td width="11%" height="1" align="center">
                      <font color="#FFFCF0">歌词</font></td>
                      <td width="12%" height="1" align="center">
                      <font color="#FFFCF0">点播</font></td>
                    </tr>
                  </table>
                  </td>
                </tr>
                <%
i=0
MaxList=10
set rs=server.createobject("adodb.recordset")
sql="SELECT * FROM MusicList where IsGood=1 order by id desc"
rs.Open sql,conn,1,1
do while not rs.eof
	i=i+1
%>
                <tr bgcolor="#FFFFFF" align="center"> 
                  <td width="5" bgcolor="#F79618" height="1">
                  　</td>
                  <td width="204" bgcolor="#FFFCF0" height="1"> <div align="left">
                    &nbsp;<b><%=i%></b>.<a href="#" onClick="window.open('YxPlay.asp?id=<%=rs("id")%>','tt90','scrollbars=no,resizable=no,width=300,height=300,menubar=no,top=168,left=168')" ><%=rs("MusicName")%></a></div></td>
                  <td width="99" bgcolor="#FFFCF0" height="1"><a href="Albumlist.asp?Name=Yxmusic&ID=<%=rs("SClassid")%>&ArtID=<%=rs("NClassid")%>"><%=rs("singer")%></a>　</td>
                  <td width="55" bgcolor="#FFFCF0" height="1"><a href="#" onClick="window.open('YxPlay.asp?id=<%=rs("id")%>','tt90','scrollbars=no,resizable=no,width=300,height=300,menubar=no,top=168,left=168')" >
                  <img border="0" src="images/real.gif"></a></td>
                  <td width="55" bgcolor="#FFFCF0" height="1">
                  <a href="http://mp3search.baidu.com/wstsearch?tn=baidump3lyric&ct=150994944&rn=10&word=<%=rs("MusicName")%>" target="_blank">
                  <img border="0" src="images/gezhi.gif" alt="查询歌词"></a></td>
                  <td width="50" bgcolor="#FFFCF0" height="1"><a href="#" onClick="window.open('Mailsong.asp?id=<%=rs("id")%>','','scrollbars=no,resizable=no,width=419,height=151,menubar=no,top=98,left=198')">
                  送友</a></td>
                </tr>
                <%
				if i>=MaxList then exit do
				rs.movenext
				loop
				rs.close
				%>
                <tr bgcolor="#FCCF95" align="center"> 
                  <td colspan="6" width="481" bgcolor="#004984" height="1"> 
                  </td>
</td>
            </tr>
          </table>
          </td>
                </tr>
              </table>
              </td>
            </tr>
            <tr>
              <td width="100%" bgcolor="#FFFFF7" valign="top" align="left" bordercolor="#F79618"><a href="http://vod.89dns.com/index.asp?dailiid=lovemx" target="_blank"><SCRIPT language=javascript>
function gowhere(formname)
{
 var url;
 if (formname.myselectvalue.value == "0")
 {
  url = "http://www1.baidu.com/baidu";
  document.search_form.tn.value = "lovemx";
  formname.method = "get";
 }
 if (formname.myselectvalue.value == "1")
 {
  url = "http://mp3search.baidu.com/wstsearch";
  document.search_form.tn.value = "baidump3";
  document.search_form.ct.value = "134217728";
  document.search_form.lm.value = "-1";
  formname.method = "get";
 }
 if (formname.myselectvalue.value == "2")
 {
  document.search_form.tn.value = "flash";
  document.search_form.ct.value = "33554432";
  url = "http://flash.baidu.com/wstsearch";
 }
 if (formname.myselectvalue.value == "3")
 { 
  document.search_form.tn.value = "baiduwstui";
  document.search_form.ct.value = "83886080";
  url = "http://www1.baidu.com/wstsearch";  
 } 
  formname.action = url;
 return true;
}
              </SCRIPT>

<table width="100%" height="60" border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#004984" width="100%" id="AutoNumber22" height="1">
<tr> 
<td><TABLE border=0 cellPadding=0 cellSpacing=0 width=100%>
<FORM name=search_form onsubmit="return gowhere(this)" target=_blank>
<INPUT name=myselectvalue type=hidden value=0>
<INPUT name=tn type=hidden>
<INPUT name=ct type=hidden>
<INPUT name=lm type=hidden>
<TBODY>
<TR> 
<TD width="28%"> <DIV align=center><A href="http://www.baidu.com/index.php?tn=lovemx" target=_blank><IMG alt=百度中文搜索引擎 border=0 height=40 src="http://img.baidu.com/images/bdlogo.gif" width=100></A></DIV></TD>
<TD vAlign=bottom width="72%"> <DIV align=center> 
<TABLE align=right border=0 cellPadding=0 cellSpacing=0 width="100%">
<TBODY>
<TR> 
<TD><FONT style="FONT-SIZE: 12px"> 
<input id=word name=word size="38" value="E线江湖"  onMouseOver="this.focus()" onblur="if (value ==''){value='E线江湖'}" onFocus="this.select()" onClick="if(this.value=='E线江湖')this.value=''">
</FONT> 
<INPUT align=middle border=0 height=23 id=image name=image src="http://img.baidu.com/img/search.gif" type=image width=51> 
<FONT style="FONT-SIZE: 12px"><A href="http://union.baidu.com/paihang/index.php" target=_blank></A></FONT></TD>
</TR>
</TBODY>
</TABLE>
</DIV></TD>
</TR>
<TR> 
<TD height=14 width="28%"> <DIV align=center><FONT style="FONT-SIZE: 12px">
<A href="http://union.baidu.com/paihang/index.php" target=_blank>
<FONT color=#ff0000>网站排行榜</FONT></A></FONT></DIV></TD>
<TD width="72%"> <DIV align=center> 
<INPUT CHECKED name=myselect onclick=javascript:this.form.myselectvalue.value=0; type=radio value=0>
<SPAN class=f12><FONT color=#0000cc style="FONT-SIZE: 12px">网页</FONT></SPAN> 
<INPUT name=myselect onclick=javascript:this.form.myselectvalue.value=1; type=radio value=1>
<SPAN class=f12><FONT color=#0000cc style="FONT-SIZE: 12px">mp3</FONT></SPAN> 
<INPUT name=myselect onclick=javascript:this.form.myselectvalue.value=2; type=radio 
value=2>
<SPAN class=f12><FONT color=#0000cc style="FONT-SIZE: 12px">flash</FONT></SPAN> 
<INPUT name=myselect onclick=javascript:this.form.myselectvalue.value=3; type=radio value=3>
<SPAN class=f12><FONT color=#0000cc style="FONT-SIZE: 12px">信息快递 </FONT></SPAN></DIV></TD>
</TR>
</FORM>
</TABLE></td>
</tr>
</table></td>
            </tr>
            <tr>
              <td width="100%" height="16" bgcolor="#FFFFFF">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
          <td width="100%" valign="bottom">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="59%">
              <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="97%">
                <tr>
              <td width="20%" align="center" height="17" valign="bottom" bgcolor="#66CCFF"><a href="new_100.asp">
              新进单曲</a></td>
              <td width="20%" align="center" height="17" valign="bottom" bgcolor="#F27C14"><a href="song100.asp">
              劲爆单曲</a></td>
              <td width="20%" align="center" height="17" valign="bottom" bgcolor="#FF0000"><a href="tuijian50.asp">
              推单曲</a></td>
              <td width="20%" align="center" height="17" valign="bottom" bgcolor="#FCCF95"><a href="newzj50.asp">
              新进专辑</a></td>
              <td width="20%" align="center" height="17" valign="bottom" bgcolor="#009ACE"><a href="zj100.asp">
              劲爆专辑</a></td>
                        </tr>
              </table>
              </td>
              <td width="41%">
              <p align="right"><img border="0" src="images/io10.gif"></td>
            </tr>
          </table>
                  </td>
                </tr>
                <tr>
          <td width="100%">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
          <td width="100%">
          <table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#004984" width="100%" id="AutoNumber34" height="1">
            <tr>
              <td width="5" bgcolor="#FF0000" height="1">　</td>
              <td width="92" bgcolor="#FFFFF7" align="center" height="1"><b>华人歌手</b></td>
              <td width="374" bgcolor="#FFFFF7" height="1">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber39">
                <tr>
                  <td width="25%" align="center">	  
                  <a class=a href="Artlist.asp?Name=Yxmusic&ID=1">
                  <font color="#000000">男歌手</font></a></td>
                  <td width="25%" align="center"> 
                  <a class=a  href="Artlist.asp?Name=Yxmusic&ID=2">
                  <font color="#000000">女歌手</font></a></td>
                  <td width="25%" align="center"> 
                  <a class=a  href="Artlist.asp?Name=Yxmusic&ID=3">
                  <font color="#000000">乐队组合</font></a></td>
                  <td width="25%" align="center"> 
                  <a class=a href="Artlist.asp?Name=Yxmusic&ID=4">
                  <font color="#000000">影视其它</font></a></td>
                </tr>
              </table>
              </td>
            </tr>
            <tr>
              <td width="5" bgcolor="#B52000" height="1">　</td>
              <td width="92" bgcolor="#FFFFF7" align="center" height="1"><b>日韩歌手</b></td>
              <td width="374" bgcolor="#FFFFF7" height="1">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber36">
                <tr>
                  <td width="25%" align="center"> 
                  <a class=a  href="Artlist.asp?Name=Yxmusic&ID=5">
                  <font color="#000000">日韩男歌手</font></a></td>
                  <td width="25%" align="center"> 
                  <a class=a href="Artlist.asp?Name=Yxmusic&ID=6">
                  <font color="#000000">日韩女歌手</font></a></td>
                  <td width="25%" align="center"> 
                  <a class=a  href="Artlist.asp?Name=Yxmusic&ID=7">
                  <font color="#000000">日韩乐队组合</font></a></td>
                  <td width="25%" align="center">　</td>
                </tr>
              </table>
              </td>
            </tr>
            <tr>
              <td width="5" bgcolor="#FF0000" height="1">　</td>
              <td width="92" bgcolor="#FFFFF7" align="center" height="1"><b>欧美歌手</b></td>
              <td width="374" bgcolor="#FFFFF7" height="1">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber37">
                <tr>
                  <td width="25%" align="center"> 
                  <a class=a  href="Artlist.asp?Name=Yxmusic&ID=9">
                  <font color="#000000">欧美歌手</font></a></td>
                  <td width="25%" align="center">　</td>
                  <td width="25%" align="center">　</td>
                  <td width="25%" align="center">　</td>
                </tr>
              </table>
              </td>
            </tr>
            <tr>
              <td width="5" bgcolor="#B52000" height="1">　</td>
              <td width="92" bgcolor="#FFFFF7" align="center" height="1"><b>其他歌手</b></td>
              <td width="374" bgcolor="#FFFFF7" height="1">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber38">
                <tr>
                  <td width="25%" align="center"> 
                  <a class=a  href="Artlist.asp?Name=Yxmusic&ID=10">
                  <font color="#000000">闽南台语</font></a></td>
                  <td width="25%" align="center"> 
                  <a class="a" href="Artlist.asp?Name=Yxmusic&ID=11">
                  <font color="#000000">儿童歌曲</font></a></td>
                  <td width="25%" align="center"> 
                  <a class=a href="Artlist.asp?Name=Yxmusic&ID=13">
                  <font color="#000000">DJ舞曲</font></a></td>
                  <td width="25%" align="center"> 
                  <a class=a href="Artlist.asp?Name=Yxmusic&ID=4">
                  <font color="#000000">影视其它</font></a></td>
                </tr>
              </table>
              </td>
            </tr>
          </table>
          </td>
                </tr>
              </table>
              </td>
                </tr>
              </table>
              </td>
            </tr>
            </table>
          </td>
        </tr>
      </table>
      </td>
      <td width="85" height="390" valign="top">
      <table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
        <tr>
          <td width="100%">
          <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="85%">
            <tr>
              <td width="100%">
                  <a href="Albumlist.asp?Name=Yxmusic&ID=2&ArtID=1325">
                  <img border="0" src="images/lx.jpg" width="70" height="414"></a></td>
            </tr>
          </table>
          </td>
        </tr>
        <tr>
          <td width="100%"><a href="#" onClick="window.open('../bbs/list.asp?boardid=5','','width=300,height=300')" >
          <img border="0" src="images/index_02.gif" width="70" height="130"></td>
        </tr>
        <tr>
          <td width="100%"><a href="#" onClick="window.open('http://www.51eline.com','','width=300,height=300')" >
                      <img border="0" src="images/index_03.gif" width="70" height="130"></td>
        </tr>
        <tr>
          <td width="100%"><a href="#" onClick="window.open('../bbs/list.asp?boardid=9','mysms','height=300,width=300,resizable=yes')" >
                      <img border="0" src="images/index_04.gif" width="70" height="130"></td>
        </tr>
      </table>
      </td>
      <td width="19" height="390" valign="top" bgcolor="#FFFFFF" background="images/bb2.gif">
      <p align="right"><img border="0" src="images/io3.gif"></td>
      <td width="174" height="390" valign="top" bgcolor="#68CCFF">
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
        <tr>
      <td width="1" height="158" valign="top">
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="1%">
        <tr>
          <td width="85%" valign="top" bgcolor="#F79618" height="1"> 
            <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber16" height="1">
              <tr>
                <td width="100%" height="1" valign="top" bgcolor="#68CCFF">
                <table border="0" cellpadding="6" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber20" height="1">
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="85%" valign="top" bgcolor="#68CCFF" height="1"> 
            <table border="0" cellpadding="6" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="1%" id="AutoNumber23" height="212">
              <tr>
                <td width="100%" bgcolor="#60CCFF">
                <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                  <tr>
                    <td width="100%" bgcolor="#60CCFF">          
                    <td width="100%" bgcolor="#60CCFF">
                    <td width="85%" valign="top" bgcolor="#60CCFF" height="1"> 

            <div align="center">
              <center>
            <table width="156" border="0" cellspacing="0" cellpadding="2" height="80" style="border-collapse: collapse" bordercolor="#111111">
<tr> 
                <td colspan="2" height="64" width="156" bgcolor="#60CCFF"> 
                    
                  <table width="156" border="0" cellspacing="0" cellpadding="2" style="border-collapse: collapse" bordercolor="#111111">
                    <tr> 
                      <td align="center" height="20">　</td>
                      <td align="center" height="20">　</td>
                      <td align="center" height="20">
<p align="left">欢迎您：<font color=#ff0000>客人</font><br>
                                视听本站的歌曲<br>
                                必须是ｅ线江湖的成员<br>
                                如果您还没有注册<br>
                                请点击以下的注册连接<br>
                                加入我们的家园<br> 
<br>
从这里进入您的[<b><a href=Box.asp?action=show target="_blank">音乐盒</a></b>]<br>
<script language=JavaScript src="../online.asp"></script>
</td>
                    </tr>
                  </table>
                  </td>
                </tr>
                <tr> 
                  
                <td width="80" height="16" bgcolor="#60CCFF">
                <p align="center">[<a href="javascript:window.close()">离开视听</a>]</td>
                  
                <td width="76" height="16" bgcolor="#60CCFF">
                [<a href="../yamen/read.htm">免费注册</a>]</td>
                </tr>
            </table>
			  </center>
            </div>
			
            </td>
                  </tr>
                </table>
                </td>
              </tr>
              <tr>
                <td width="100%" bgcolor="#60CCFF">
                <p align="center">
                <a href="http://windowsmedia.com/download" target="_blank">
                <img border="0" src="images/WMP9series_Free.gif" alt="如不能正常播放请下载此Media player 9播放器"></a></td>
              </tr>
              <tr>
                <td width="100%" bgcolor="#60CCFF">
                <p align="center">如不能正常播放请下载此<br>
                Media player 9播放器<br><a href="myasp.asp" target="_blank">[<font color="#FF00ff">测试你的上网速度</font>]</a></td>  
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="85%" valign="top" bgcolor="#68CCFF" height="1"> 
            <img border="0" src="images/io6.gif"></td>
        </tr>
        <tr>
          <td width="85%" valign="top" bgcolor="#68CCFF" height="1"> 
            <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber17" cellspacing="6">
              <tr>
                <td width="20%"><%
m=0
set rs=conn.execute("SELECT top 10 id,Singer,ClassID,SClassID,NClassID,MusicName,hits FROM MusicList order by hits desc") 
if not Rs.eof then
do while not rs.eof
m=m+1
%>
                    <tr> 
                      <td width="8%" align="center" valign="top">
                      <img border="0" src="images/img_arrow02.gif"></td>
                      <td width="147%" valign="top"><a href="#" onClick="window.open('YxPlay.asp?id=<%=rs("id")%>','Ting88','scrollbars=no,resizable=no,width=300,height=300,menubar=no,top=168,left=168')" ><%=rs("musicName")%></a></td>
                      <td width="17%" valign="top"><%=rs("hits")%></td>
                    </tr>
                    <%
if m>=10 then exit do
rs.movenext 
loop
else
response.write "尚无收录"
end if
rs.close
%></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="85%" valign="top" bgcolor="#68CCFF" height="1"> 
            <img border="0" src="images/io7.gif"></td>
        </tr>
        <tr>
          <td width="85%" valign="top" bgcolor="#68CCFF" height="1"> 
            <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber18">
              <tr>
                <td width="100%">
                <table border="0" cellpadding="0" cellspacing="6" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber19">
                  <tr>
                    <td width="16%"><%
					m=0
					set rs=conn.execute("SELECT top 10 Specialid,name,Classid,SClassid,NClassid,NClass,hits FROM Special order by hits desc") 
					if not Rs.eof then
					do while not rs.eof
					m=m+1
					%>
                    <tr> 
                      <td width="1%" align="center">
                      <img border="0" src="images/img_arrow02.gif"></td>
                      <td width="157%"><a href='MusicList.asp?AlbumID=<%=rs("Specialid")%>'><%=rs("name")%></a></td>
                      <td width="12%"><%=rs("hits")%></td>
                    </tr>
                    <%
					if m>=10 then exit do
					rs.movenext 
					loop
					else
					response.write "尚无收录"
					end if
					rs.close
					%>
</td>
                  </tr>
                </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="85%" valign="top" bgcolor="#68CCFF" height="1"> 
            <p align="center"></td>
        </tr>
        <tr>
</td>
        </tr>
      </table>
      </td>
        </tr>
      </table>
      </td>
    </tr>
    </table>
  </center>
</div>

</body>
</html>
<!--#include file="Bottom.asp"-->
