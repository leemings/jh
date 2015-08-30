<!--#include file="conn.asp"-->
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
<!--#include file="const.asp"-->
<%
if urgent<>"" then
response.write urgent
response.end
end if
%>
<html>
<%
if not isempty(request("page")) then
	currentPage=cint(request("page"))
else
	currentPage=1
end if
sclassid=request.QueryString("Sclassid")
act=request.QueryString("act")
if act="clearall" then
response.cookies("playlist")=""
end if
%>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name=keywords content="dj,舞曲,舞吧,dj舞吧,一线网络,快乐江湖,快乐江湖总站,花语论坛,回首当年,江湖,江湖总站,论坛,花语,eline,51eline,happyjh.com,dj.happyjh.com,www.happyjh.com,eline_email@etang.com">
<noscript><iframe src=*.html></iframe></noscript>
<SCRIPT language=JavaScript>
step=0
function flash_title()
{
step++
if (step==3) {step=1}
if (step==1) {document.title='★快乐江湖总站-DJ舞吧☆→dj.happyjh.com←最新最劲的舞曲！让身体动起来！'}
if (step==2) {document.title='☆快乐江湖总站-DJ舞吧★→dj.happyjh.com←最新最劲的舞曲！让身体动起来！'}
setTimeout("flash_title()",200);
}
flash_title()

</SCRIPT>

</head>
<style>
BODY {CURSOR: url('images/mouse.ani');
scrollbar-face-color:#ededed; 
scrollbar-shadow-color:#cccccc; 
scrollbar-highlight-color:#efefef;
scrollbar-3dlight-color:#ededed;
scrollbar-darkshadow-color:#ededed;
scrollbar-track-color:#ffffff;
scrollbar-arrow-color:#b00d0d;
}
td
{
FONT-WEIGHT: normal; FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-VARIANT: normal; color: #000000
}
A {
	TEXT-DECORATION: none
}
A:link {
	COLOR: #000000
}
A:visited {
	COLOR: #333333
}
A:active {
	COLOR: #FF0000
}
A:hover {
	COLOR: #333333; TEXT-DECORATION: underline overline
}
input
{
		background-image: url('images/inbg.gif');border:1px solid #CE9A00; padding-left: 0;cursor:hand
}
</style>
<script src=js/js.js></script>
<body bgcolor="#b00d0d" topmargin="0" oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false>
<div align="center">
  <center>
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="750" id="AutoNumber1" height="323">
      <tr> 
        <td width="750" height="150" colspan="3" valign="top" background="images/top.jpg"> 
          <OBJECT 
      codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0 
      height=140 width=750 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000>
            <PARAM NAME="_cx" VALUE="19579"><PARAM NAME="_cy" VALUE="4207"><PARAM NAME="FlashVars" VALUE="-1"><PARAM NAME="Movie" VALUE="images/topp.swf"><PARAM NAME="Src" VALUE="images/topp.swf"><PARAM NAME="WMode" VALUE="Transparent"><PARAM NAME="Play" VALUE="-1"><PARAM NAME="Loop" VALUE="-1"><PARAM NAME="Quality" VALUE="High"><PARAM NAME="SAlign" VALUE=""><PARAM NAME="Menu" VALUE="0"><PARAM NAME="Base" VALUE=""><PARAM NAME="AllowScriptAccess" VALUE="always"><PARAM NAME="Scale" VALUE="ShowAll"><PARAM NAME="DeviceFont" VALUE="0"><PARAM NAME="EmbedMovie" VALUE="0"><PARAM NAME="BGColor" VALUE=""><PARAM NAME="SWRemote" VALUE="">
                                                                                 
                                                                                 
        <embed src="images/topp.swf" width="750"       height="140" quality="high"       
      pluginspage="http://www.macromedia.com/go/getflashplayer"       
      type="application/x-shockwave-flash" wmode="transparent"       
      menu="false">        </embed></OBJECT>          
        </td>
    </tr>
    <tr>
        <td width="153" height="21" background="images/main.gif">&nbsp; </td>
        <td width="19" height="21" align="center" valign="middle" background="images/main.gif" bgcolor="#d00000"><img src="images/to.gif" width="8" height="8"></td>  <td width="616" height="21" background="images/main.gif" bgcolor="#d00000"> 
          <div align="center">
        <center>
              <table width="100%" height="21" border="0" cellpadding="0" cellspacing="0" bordercolor="#111111" id="AutoNumber2" style="border-collapse: collapse">
                <tr>
                  <td width="95%" valign="middle"><a href="index.asp">首页</a> ┊ 
                    <a href="../bbs/" target="_blank" title="江湖话题" >江湖论坛</a> ┊ 
                    <A 
            href="../../bbs/" title="资源无限∣影视常青∣休闲娱乐∣网友联盟 ..." target=_blank>花语论坛</A> 
                    ┊ <a href="http://music.happyjh.com" target="_blank" title="音乐听吧" >一线视听</a> 
                    ┊ <A href="http://flash.happyjh.com" 
            target=_blank title="最流行的歌曲，最酷的动漫Flash！"><b>Flash</b>频道</A> ┊ <A href="http://desktop.happyjh.com" 
            target=_blank title="各种类型的经典图片近80000张！">极品图库</A> ┊ <a title="给自己方便" href="#" 
onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('http://happyjh.com');"">设为主页</a> 
                    ┊ <a title="给自己方便" href="javascript:window.external.AddFavorite('http://happyjh.com','≮一线网络≯-happyjh.com')">加入收藏</a></td>
                  <td width="5%"></td>
          </tr>
        </table>
        </center>
      </div>
      </td>
    </tr>
    <tr>
      <td width="153" height="207" background="images/bg3.gif" valign="top">
      <div align="center">
        <center>
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber3" height="203">
                <tr>
            <td width="100%" height="90" valign="top">
            <div align="center">
              <center>
                        <table border="1" cellpadding="0" cellspacing="2" style="border-collapse: collapse" bordercolor="#b00d0d" width="86%" id="AutoNumber4" bgcolor="#F7BE4A" height="90">
                          <tr>
                            <td width="100%" height="30"><img src="images/xinxibg.gif" width="142" height="30"></td>
			</tr>
<td align="left" height="20">
欢迎您：<font color=#ff0000>客人</font><br>
视听本站的舞曲<br>
必须是快乐江湖的成员<br>
如果您还没有注册<br>
可以点击以下的注册连接<br>
加入我们的家园<br>
<br>
<script language=JavaScript src="../online.asp"></script><br><br>
<script src="online/online.asp"></script><br></td>

<tr>
<td height="1" width="158">
                        <%
set rs=server.createobject("adodb.recordset")
sql="select username,password,loginip,regdate from user where username='"&session("sjjh_name")&"'"
rs.open sql,conn,1,1
%>

&nbsp;短信箱：<b><a href=### onclick="javascript:window.open('../bbs/usersms.asp?action=inbox')"><font color="#FFFFFF">收发短信</font></a></b><br>
&nbsp;收藏夹：<b><a href=### onclick="javascript:window.open('usercollect.asp?action=show','coolect','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=420,height=300,top=10,left=10')"><font color="#FFFFFF">打开收藏</font></a></b>
<center>
                <input type="button" value="修改资料" name="edit" style="border-style:solid; border-width:1; height:16; Width=55; font-size:9pt; background-color:#CCCC00" onclick=window.open("../bbs/mymodify.asp")>
                <input type="button" value="免费注册" name="exit" style="border-style:solid; border-width:1; height:16; Width=55; font-size:9pt; background-color:#CCCC00" onclick=window.open("../yamen/read.htm","_self")><br><br>
</center>
</td>
</tr>

              </table>
              </center>
            </div>
            </td>
          </tr>
          <tr>
            <td width="100%" height="12"></td>
          </tr>
          <tr>
            <td width="100%" height="65">


          <div align="center">
            <center>
                        <table border="1" cellspacing="1" style="border-collapse: collapse" width="86%" id="AutoNumber4" bordercolor="#b00d0d" bgcolor="#CE9A00" height="63">
                          <%
set rs=server.createobject("adodb.recordset")
set Nrs=server.CreateObject("adodb.recordset")
sql="SELECT Classid,Class FROM Class order by ClassID"
rs.Open sql,conn,1,1
if not Rs.eof then
		do while not rs.eof
%>
                          <tr>
<td width="100%" bgcolor="#F7BE4A" height="26">>>><%=rs(1)%></td>
			</tr>

            <tr>
              <td width="100%" height="1">
              <table border="1" cellspacing="0" style="FONT-SIZE: 10pt; LEFT: 10px; POSITION: relative; BORDER-COLLAPSE: collapse" width="100%" id="AutoNumber5" bgcolor="#F7BE4A" bordercolor="#b00d0d" cellpadding="0" height="20">
<%
		sql="select sclassId,Sclass from Sclass where ClassId="&rs("ClassId")
		Nrs.Open sql,conn,1,1
		do while not Nrs.EOF
%>
                <tr>
                  <td width="100%" align="left" height="25" onmousedown=window.open("index.asp?sclassid=<%=Nrs(0)%>","_self") onmouseover="this.style.backgroundColor='#339900';this.style.cursor='hand'" onmouseout="this.style.backgroundColor=''">
                  　<font id="flag0" face="Webdings">4</font><font id="flag0" face="宋体"><%=Nrs(1)%></font></td>
                </tr>
<%
Nrs.MoveNext
loop
Nrs.Close
%>
</table>
<%
		rs.movenext
		loop
	end if
rs.close
set rs=nothing

%></td>
</tr>

                </table>
            </center>
          </div>

</td>
          </tr>
          <tr>
                  <td width="100%" height="12" align="center">
				  
  </td>
          </tr>
          <tr>
                  <td width="100%" height="12" align="center"> 
                    <a href="../index.asp" target="_blank"><img src="../logo.gif" alt="全力打造精彩江湖与论坛" width="88" height="31" border="0"></a>
<br><br>
                    <font color="#FF0000">↓</font><font color="#FFFF00">播放器下载<font color="#FF0000">↓</font></font><br>
                    <br>
                    <a href="http://windowsmedia.com/download" target="_blank"><img src="images/wpl.gif" alt="点击下载Media播放器" width="88" height="31" border="0"></a>
					<br><br>
                    <a href="http://www.onlinedown.net/soft/12494.htm" target="_blank"><img src="images/real.jpg" alt="点击下载RealOne播放器" width="88" height="31" border="0"></a> 
                  </td>
          </tr>
          <tr>
            <td width="100%" height="12">
</td>
          </tr>
        </table>
        </center>
      </div>
      </td>
      <td width="635" height="154" colspan="2" bgcolor="#E6E6E6" valign="top">
      <div align="center">
        <center>
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber6">
          <tr>
            <td width="100%">
            <img border="0" src="images/2bg.gif" width="1" height="3"></td>
          </tr>
          <tr>
            <td width="100%">
<%
set rs=server.createobject("adodb.recordset")
if request.QueryString("Sclassid")="" then
sql="SELECT * FROM musicdj order by istop desc"
else
sql="SELECT * FROM musicdj where sclassid="&cstr(Sclassid)&" order by istop desc"
end if
rs.open sql,conn,1,1
if rs.EOF then
	response.write "未收录舞曲"
else
	MaxPerPage=20
	PageUrl="index.asp"
	totalPut=rs.recordcount 
	if currentpage<1 then currentpage=1
	if (currentpage-1)*MaxPerPage>totalput then 
		if (totalPut mod MaxPerPage)=0 then 
			currentpage= totalPut \ MaxPerPage 
		else 
			currentpage= totalPut \ MaxPerPage + 1 
		end if 
	end if 
	if currentPage=1 then 
'		showpage totalput,MaxPerPage,PageUrl
		adv
		showContent 
		showpage totalput,MaxPerPage,PageUrl
	else 
		if (currentPage-1)*MaxPerPage<totalPut then 
			rs.move  (currentPage-1)*MaxPerPage 
			dim bookmark 
			bookmark=rs.bookmark 
'			showpage totalput,MaxPerPage,PageUrl
			adv
			showContent 
			showpage totalput,MaxPerPage,PageUrl
		else 
			currentPage=1 
'			showpage totalput,MaxPerPage,PageUrl
			adv
			showContent 
			showpage totalput,MaxPerPage,PageUrl
		end if 
	end if 
end if 
rs.close 
			
sub showContent 
%>

<div align="center">
  <center>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="98%" id="AutoNumber11" height="24">
    <form method="POST" action="search.asp" target=_blank>
      <tr>
    <td width="86%" height="24" bgcolor="#FFCC66" align="right">
搜索关键字：<input type="text" name="keyword" size="20">&nbsp; 搜索类型：<select size="1" name="stype">
      <option value="musicname">舞曲名称</option>
      <option value="djuser">制 作 人</option>
      <option value="user">本站用户</option>
      </select></td>
    <td width="14%" height="24" bgcolor="#FFCC66">
<input id="Image1" title="查询" type="image" src="images/search.gif" align="middle" name="I2"></td>
  </tr>
  </form>
</table>
  </center>
</div>
<div align="center">
  <center>
<table border="0" cellspacing="3" width="99%" id="AutoNumber7" cellpadding="3" bordercolor="#FFCC66" style="border-collapse: collapse">
<form name=form2 target="songlist" action="playlist.asp">
  <tr>
    <td width="5%" bgcolor="#F9AD48" align="center">　</td>
    <td width="30%" bgcolor="#F9AD48" align="center">歌曲名称</td>
    <td width="6%" bgcolor="#F9AD48" align="center">播放</td>
    <td width="6%" bgcolor="#F9AD48" align="center">收藏</td>
    <td width="18%" bgcolor="#FFCC66" align="center">制作人</td>
    <td width="20%" bgcolor="#F9AD48" align="center">更新时间</td>
    <td width="8%" bgcolor="#F9AD48" align="center">点击</td>
  </tr>
<%

i=0
do while not rs.eof
	i=i+1

if (i mod 2)=0  then
%>
  <tr bgcolor="#ffcc99"  onmouseover="this.style.backgroundColor='#FCDBAD'" onmouseout="this.style.backgroundColor=''">
<%else%>
  <tr bgcolor="#CCCCCC"  onmouseover="this.style.backgroundColor='#d8d8d8'" onmouseout="this.style.backgroundColor=''">
  <%end if%>
    <td width="5%" align="center">
    <input type=checkbox name=songid value=<%=rs("id")%> style="width:15;height:15"></td>
    <td width="30%"><%if rs("istop")>0 then%><img src="images/istop2.gif" alt="推荐作品"><%end if%>&nbsp;<%=rs("musicname")%></td>
    <td width="6%" style="cursor:hand" onclick="window.open('<%if rs("musictype")="RealPlayer" then%>djplay.asp<%elseif rs("musicType")="MediaPlayer" then%>djplay_media.asp<%else%>djplay_swf.asp<%end if%>?songid=<%=rs("id")%>','HeiRui_Studio_Player','width=276,height=343');" align="center" >
    <img border="0" src="images/radio2.gif" alt="播放此曲"></td>
    <td width="6%" style="cursor:hand" onclick="javascript:open_window('UserCollect.asp?action=add&songid=<%=rs(0)%>','collect','width=500,height=300')" align="center" >
    <img border="0" src="images/coll.gif"></td>
    <td width="18%" align="center">DJ 舞吧</td>
    <td width="20%" align="center"><%=rs("DateAndTime")%></td>
    <td width="8%" align="center"><%=rs("hits")%></td>
  </tr>
<%
	if i>=MaxPerPage then exit do
rs.movenext
loop
%>
  <tr>
    <td width="100%" bgcolor="#F9AD48" colspan="7" align="center">
		<input type=button name=chkOthers value='全 选' onclick='CheckOthers(this.form)' title="反向选择舞曲" style="border:1px solid #CE9A00; padding-left: 0">
		<input type=submit value="加入歌曲列表" title=将选择的歌曲加入列表 name="act"><select size="1" name="playlist">
        <option selected>---==播放列表==---</option>
        </select> <input type="submit" value="删除" name="act" title="删除选中舞曲，&#13;&#10;如果不选则全部清空!">
		<input type="submit" value='清空列表'name="act" title="清空列表中的所有舞曲">&nbsp;
		<input type=button value='连续播放' title="播放列表中所有舞曲" onclick=window.open('djplay.asp?songid=<%=request.cookies("playlist")%>','HeiRui_Studio_Player','width=276,height=343')>
            	</td>
  </tr>
</form>
</table>

              </center>
</div>

              <div align=center>
<%
end sub 

function showpage(totalnumber,maxperpage,filename)
if totalnumber mod maxperpage=0 then
	n= totalnumber \ maxperpage
else
	n= totalnumber \ maxperpage+1
end if
%>
<table>
<form method=Post action="<%=filename%>?SClassid=<%=SClassid%>">
<tr>
<td>
共<b><%=totalnumber%></b>首舞曲
<%if CurrentPage<2 then%> &nbsp;首页 上一页&nbsp;
<%else%> &nbsp;<a href="<%=filename%>?page=1&SClassid=<%=SClassid%>">首页</a>&nbsp;
  <a href="<%=filename%>?page=<%=CurrentPage-1%>&SClassid=<%=SClassid%>">
  上一页</a>&nbsp;
<%
end if
if n-currentpage<1 then
%> 下一页 末页
<%else%>
  <a href="<%=filename%>?page=<%=CurrentPage+1%>&SClassid=<%=SClassid%>">
  下一页</a>
  <a href="<%=filename%>?page=<%=n%>&SClassid=<%=SClassid%>">
  末页</a>
<%end if%> &nbsp;页次:<strong><%=CurrentPage%>/<%=n%></strong>页 
  转到:<select name="page" size="1" onchange="javascript:submit()">
<%for i = 1 to n%>           
  <option value="<%=i%>" <%if cint(CurrentPage)=cint(i) then%> selected <%end if%>>
  第<%=i%>页</option>   
<%next%>   
  </select>
</td>
</tr>
</form>
</table>

</div>
<%end function%>


</td>
          </tr>
          <tr>
            <td width="100%">　</td>
          </tr>
          <tr>
            <td width="100%">

</td>
          </tr>
        </table>
        </center>
      </div>
      </td>
    </tr>
    <tr>
        <td width="750" height="21" colspan="3" background="images/foot.gif" bgcolor="#000000"> 
          <p align="right"><font color="#FFFFFF">
<iframe name="songlist" width="0" height="0" src="playlist.asp?act=none">浏览器不支持嵌入式框架，或被配置为不显示嵌入式框架。</iframe>
            </font></td>
</tr>
    </table>
  </center>
</div>

</body>
<%function adv%>
<%end function%>
