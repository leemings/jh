<!-- ＤＶＢＢＳ６．０论坛首页调用测试与帮助 Script Written by fssunwin -->
<%
dim bbsurl
bbsurl="http://www.artbbs.net/"
%>
<style type="text/css">
A:link,A:active,A:visited{TEXT-DECORATION:none ;Color:#000000}
A:hover{TEXT-DECORATION: underline;Color:#4455aa;
}
BODY{
FONT-SIZE: 12px;
COLOR: #000000;
font-family:Verdana,arial,Tahoma,宋体;
background-color: #ffffff; }
TD{
font-family: Verdana,arial,Tahoma,宋体;
font-size: 12px;
line-height: 15px;
}
</style>
<CENTER>
<FONT style="FONT-SIZE: 25px;line-height: 30px;" face=Verdana size=4>
<b>ＤＶＢＢＳ６．０论坛首页调用测试与帮助</b></font>
</CENTER>
<BR><BR><BR>

基本操作步骤:<BR><font color=red>
１：上传到与论坛ＣＯＮＮ．ＡＳＰ文件相同的目录下；<BR>
２：用记事本编辑newtopic.asp，tongji.asp文件，修改论坛地址（该文件里有注明）；<BR>
３：在需要调用的网页文件中加入调用代码，如：&lt;script src=ＸＸＸＸ&gt;&lt;/script></font><BR>   
<BR>
<font color="#FF0000"><b>帖子数据调用：newtopic.asp调用演示,论坛帖子调用</b></font>
<P>
相关参数<br>
boardid :版面id，全部为all<br>
bname&nbsp;&nbsp; :0:为不调用  1:调用论坛名称<br>   
tlen&nbsp;&nbsp;&nbsp; :标题长度<br>
n　　 　:显示多少个标题<br>
sdate 　:查询多少天内帖子，1为当天(若为空则日期不限,建议为空)<br>
orders　: 排序方法，1为按照点击(最热帖)，2为按照时间(按最新回复时间),3:为按照时间(按最新主题时间)<br>
info&nbsp;&nbsp;&nbsp; :1为显示发表时间和用户，2为显示发表时间，3为显示发表用户，4为显示发表用户和点击数，5为显示点击数，6为显示发表日期和用户，7为显示发表日期，0为不显示<br>
action&nbsp; :1: 显示主题  2:显示精华主题 3:显示主题或回复<BR>   
reply&nbsp;&nbsp; :1:显示主题的回复时间及显示回复人姓名  0:显示主题的作者及发表时间<br>
lock&nbsp;&nbsp;&nbsp; :0 关闭限制调用 1:打开限制调用(同时修改newtopic.asp文件里限制版块的参数,内附说明)<br>
showpic&nbsp;&nbsp;:0 关闭心情图片 ; 1 显示帖子心情图片
<br>
<p>
演示１：在页面插入的代码（以下为带论坛名称的主题调用）:<p><font color="#FF0000"><span style="background-color: #FFFFCC">&lt;script src=<%=BBSURL%>newtopic.asp?boardid=all&amp;lock=0&amp;bname=1&amp;tlen=16&amp;n=10&amp;sdate=&amp;orders=2&amp;info=6&amp;action=1&amp;reply=0&amp;showpic=0&gt;&lt;/script></span></font><BR>
<BR>
<script src=newtopic.asp?boardid=all&lock=0&bname=1&tlen=16&n=10&sdate=&orders=2&info=6&action=1&reply=0&showpic=0>
</script>
<BR>
演示２：在页面插入的代码（以下为不带论坛名称,显示心情图标,只显示发表时间的主题调用）:<p><font color="#FF0000"><span style="background-color: #FFFFCC">&lt;script src=<%=BBSURL%>newtopic.asp?boardid=all&amp;lock=0&amp;bname=0&amp;tlen=16&amp;n=10&amp;sdate=&amp;orders=2&amp;info=7&amp;action=1&amp;reply=0&amp;showpic=1&gt;&lt;/script></span></font><BR>
<script src=newtopic.asp?boardid=all&lock=0&bname=0&tlen=16&n=10&sdate=&orders=2&info=7&action=1&reply=0&showpic=1>
</script>
<BR>
============================================================================</p>
<p align="left"><b><font color="#FF0000">论坛展区：newsfile.asp文件的调用方法</font></b><BR><BR>
相关参数:<BR>
type：调用文件类型，默认为all。（参数设置：1=图片集,2=FLASH集,3=音乐集,4=电影集,0=文件集）<BR>
boardid：调用版块ＩＤ，默认为all。<BR>
lock：分版调用定制开关。（参数设置：0=关闭，1=不被调用，2=只允许调用）；当boardid=all,此项设置才能生效，受定制的版块ＩＤ请在newsfile.asp文件中设置。<BR>
orders：排序方式，1为按照点击浏览（最热），2为按照新增时间（最新）<BR>
topic：显示文字主题。（参数设置：0=关闭，1=显示）<BR>
n：显示调用的文件数<BR>
tab：每行显示个数<BR>
mode：调用模式::0=平面模式,1=向下滚动模式,2=向上滚动模式,3=向左滚动模式,4=向右滚动模式<BR>
<b>另外</b>：表格颜色，图片显示大小等请newsfile.asp文件中设置，详细看文件内附说明。
<BR>
<p><font color="#FF0000"><span style="background-color: #FFFFCC">&lt;script src=<%=BBSURL%>newsfile.asp?type=1&amp;boardid=all&amp;lock=0&amp;orders=2&amp;topic=0&amp;n=8&amp;tab=4&amp;mode=0>&lt;/script></span></font></p>
<script src=newsfile.asp?type=1&boardid=all&lock=0&orders=2&topic=0&n=8&tab=4&mode=0></script>
<br>

  <table border=0 style="background-color: #000000" cellspacing=1 cellpadding=5 align=center>
    <tr>
      <td height="20" colspan="4" bgcolor="#999999" align=center><b><font color=white>四种滚动模式例子(以下均为调用8条信息: n=8)</font></b></td>
    </tr>
    <tr>
      <td height="20" bgcolor="#E3E3E3"><font color="#FF0000">①</font>参数:tab=1:mode=1</td>
      <td bgcolor="#E3E3E3"><font color="#FF0000">②</font>参数:tab=1:mode=2</td>
      <td bgcolor="#E3E3E3"><font color="#FF0000">③</font>参数:tab=8:mode=3</td>
      <td bgcolor="#E3E3E3"><font color="#FF0000">④</font>参数:tab=8:mode=4</td>
    </tr>
    <tr>
      <td height="100" align=center><script src=newsfile.asp?type=1&boardid=all&lock=0&orders=2&topic=1&n=8&tab=1&mode=1></script></td>
      <td align=center bgcolor="#808080"><script src=newsfile.asp?type=1&boardid=all&lock=0&orders=2&topic=1&n=8&tab=1&mode=2></script></td>
      <td align=center><script src=newsfile.asp?type=1&boardid=all&lock=0&orders=2&topic=1&n=8&tab=8&mode=3></script></td>
      <td align=center bgcolor="#808080"><script src=newsfile.asp?type=1&boardid=all&lock=0&orders=2&topic=1&n=8&tab=8&mode=4></script></td>
    </tr>
	<tr>
      <td colspan="4" bgcolor="#E3E3E3">四种模式分别代码参考:<BR>
	  <font color="#FF0000">①:</font><font color="#ff0000"><span style="background-color: #ffffcc">&lt;script src=newsfile.asp?type=1&amp;boardid=all&amp;lock=0&amp;orders=2&amp;topic=1&amp;n=8&amp;tab=1&amp;mode=1>&lt;/script></span></font><BR>
	  <font color="#FF0000" >②:</font><font color="#ff0000"><span style="background-color: #ffffcc">&lt;script src=newsfile.asp?type=1&amp;boardid=all&amp;lock=0&amp;orders=2&amp;topic=1&amp;n=8&amp;tab=1&amp;mode=2>&lt;/script></span></font><BR>
	  <font color="#FF0000" >③:</font><font color="#ff0000"><span style="background-color: #ffffcc">&lt;script src=newsfile.asp?type=1&amp;boardid=all&amp;lock=0&amp;orders=2&amp;topic=1&amp;n=8&amp;tab=8&amp;mode=3>&lt;/script></span></font><BR>
	  <font color="#FF0000" >④:</font><font color="#ff0000"><span style="background-color: #ffffcc">&lt;script src=newsfile.asp?type=1&amp;boardid=all&amp;lock=0&amp;orders=2&amp;topic=1&amp;n=8&amp;tab=8&amp;mode=4>&lt;/script></span></font>
	</td>
    </tr>
  </table>
<BR>
============================================================================</p>
<p align="left"><b><font color="#FF0000">论坛统计信息：tongji.asp文件的调用方法</font></b><BR><BR>
相关参数:<BR>
1: orders=1 论坛统计&nbsp;<br>
2: orders=2 发帖TOP用户排行&nbsp;<br>
3: orders=3 新注册用户排行<br>
4: orders=4 论坛版块调用<BR>
5: orders=5 论坛公告调用<BR>
<br>
<b>(一) 论坛统计</b> 
<BR>
在页面插入的代码:</p>
<p><font color="#FF0000"><span style="background-color: #FFFFCC">&lt;script src=<%=BBSURL%>tongji.asp?orders=1&gt;&lt;/script></span></font></p>

<script src=tongji.asp?orders=1></script>

<br><br>
<b>(二) 发帖前10名会员</b>
<BR>
n：显示多少个名<br>
在页面插入的代码:<p><font color="#FF0000"><span style="background-color: #FFFFCC">&lt;script src=<%=BBSURL%>tongji.asp?orders=2&amp;n=10>&lt;/script></span></font></p>
<script src=tongji.asp?orders=2&n=10></script>
<br>
<br>
<b>(三) 新注册的10名会员</b>   
<BR>
n：显示多少个名<br>
在页面插入的代码:<p><font color="#FF0000"><span style="background-color: #FFFFCC">&lt;script src=<%=BBSURL%>tongji.asp?orders=3&amp;n=10>&lt;/script></span></font></p>
<script src=tongji.asp?orders=3&n=10></script>
<br>
<p>
<b>(四) 论坛版块调用</b>
<p>参数说明:</p>
<p>
boardid: =all 调用所有版块;=0 调用是上一级分类版块 ;=选择要显示下级分类版块的该版ID<br>
stats&nbsp; : =0 显示所有版块;=1 不显示隐藏版块<br>
depth: 限制调用版块的层数(如0,表示只调用每一层分类)<br>
model: 版块调用的模式(=0 树型结构 ;=1 地图结构)<br>
<p>演示1(model=0):树型结构<br>
<font color="#FF0000"><span style="background-color: #FFFFCC">&lt;script src=<%=BBSURL%>tongji.asp?orders=4&amp;boardid=all&amp;stats=1&amp;depth=1&amp;model=0>&lt;/script></span></font><br>
<br><script src=tongji.asp?orders=4&boardid=all&stats=1&depth=1&model=0></script>
<p>
演示2(model=1):地图结构<br>
<font color="#FF0000"><span style="background-color: #FFFFCC">&lt;script src=<%=BBSURL%>tongji.asp?orders=4&amp;boardid=all&amp;stats=1&amp;depth=1&amp;model=1>&lt;/script></span></font><br>
<script src=tongji.asp?orders=4&boardid=all&stats=1&depth=1&model=1></script>

<br>
<p>
<b>(五) 论坛公告调用</b>
<BR>
<p>参数说明:
boardid: =all 调用所有版块;=0 表示首页公告 ;=该版ID的公告<br>
model: =1 换行模式 ; =2 滚动模式<br>
n: 显示的数量<br>
tlen: 标题长度,可以为空<br>
<p>演示1(model=1):换行模式<br>
在页面插入的代码:<p><font color="#FF0000"><span style="background-color: #FFFFCC">&lt;script src=<%=BBSURL%>tongji.asp?orders=5&amp;boardid=all&amp;model=1&amp;n=5&amp;tlen=16&gt;&lt;/script></span></font></p>
<script src=tongji.asp?orders=5&boardid=all&model=1&n=5&tlen=16></script>
<p>
演示2(model=2):滚动模式<br>
在页面插入的代码:<p><font color="#FF0000"><span style="background-color: #FFFFCC">&lt;script src=<%=BBSURL%>tongji.asp?orders=5&amp;boardid=all&amp;model=2&amp;n=5&amp;tlen=&gt;&lt;/script></span></font></p>
<script src=tongji.asp?orders=5&boardid=all&model=2&n=5&tlen=></script>