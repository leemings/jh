<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
call nav()
stats="在线测试"
call head_var(2,0,"","")
if Cint(GroupSetting(14))=0 then
Errmsg=Errmsg+"<br>"+"<li>您没有在本社区在线测试的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
call dvbbs_error()
else
response.write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1><tr><th valign=middle colspan=2 align=center height=25><b>在线测试</b></td></tr><tr><td valign=middle class=tablebody1 height=100><CENTER>"
'本程序版权属于晨阳在线网站所有
%>

<html>

<head>
<title>在线算命</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<style>
<!--
BODY {SCROLLBAR-FACE-COLOR: #D6D6D6; SCROLLBAR-HIGHLIGHT-COLOR: #3A6EA5; SCROLLBAR-SHADOW-COLOR: #ffffff; SCROLLBAR-3DLIGHT-COLOR: #FFFFFF; SCROLLBAR-ARROW-COLOR:  #8FA5B6; SCROLLBAR-TRACK-COLOR: #f3f3f3; SCROLLBAR-DARKSHADOW-COLOR: #3A6EA5; }
-->
</style>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="2" topmargin="2">
<table width="600" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="600" height="23" background="ontest/bg1.gif">
    <table border="0" cellspacing="0" width="600" id="AutoNumber2" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
      <tr>
        <td width="25">
        <img border="0" src="ontest/err1.gif" width="24" height="23"></td>
        <td width="536" align="left">
        <table border="0" cellspacing="0" width="40%" id="AutoNumber5" cellpadding="0" height="15">
          <tr>
            <td width="100%" valign="bottom" height="15">&nbsp;&nbsp;
            <font color="#FF0000">&nbsp;爱情、人际、个性</font><font color="#0000FF"> 
            在线测试：</font></td>
          </tr>
        </table>
        </td>
        <td width="39" valign="bottom" style="border-right: 1px solid #000000">
        <img border="0" src="ontest/help.gif" start="fileopen" width="15" height="18" align="absbottom">
        <a href="javascript:window.close()">
        <img border="0" src="ontest/close.gif" width="15" height="18" align="absbottom" alt="关闭窗口"></a></td>
      </tr>
    </table>
    </td>
  </tr>
  <tr>
    <td height="562" valign="top" style="border-left: 1px solid #000000; border-right: 1px solid #000000; border-bottom: 1px solid #000000">
    <div align="center">
      <center>
      <table border="1" cellspacing="1" height="570" style="border-left: 1px solid #000000; border-right: 1px solid #000000; border-bottom: 1px solid #000000" width="596" id="AutoNumber3" bgcolor="#ECEDEF">
        <tr>
          <td width="100%"> 
          <div align="center">
            <center>
            <table border="1" width="580" id="AutoNumber6" height="549" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" cellspacing="1" background="ontest/bg2.gif">
              <tr>
                <td width="80" height="145" valign="top" align="center">
                <p style="margin-top: 5">
                <img border="0" src="ontest/love.GIF" width="80" height="60"></p>
                <p><font color="#0000FF">[</font> <font color="#FF0000">爱情测试</font>
                <font color="#0000FF">]</font>　</td>
                <td width="500" height="145" valign="top">
                <div align="center">
                  <center>
                  <table border="0" cellspacing="3" width="95%" id="AutoNumber7" cellpadding="0" height="1" background="ontest/bg3.gif">
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=001>速算恋爱成功率</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=002>你是否感情用事</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=003>性成熟程度测试(男性)</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=004>性成熟程度测试(女性)</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=005>爱情方向自测_初恋时光</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=006>合格丈夫测试</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=007>婚约危机</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=008>何种异性将对你倾心</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=009>爱情的失误</td>
                    </tr>
                    <tr>
                      <td width="33%" height="17"><a href=z_answer.asp?id=010>最适合你的示爱方式</a></td>
                      <td width="33%" height="17"><a href=z_answer.asp?id=011>谁令你一见倾情?</a></td>
                      <td width="34%" height="17"><a href=z_answer.asp?id=012>测试你的性取向</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="17"><a href=z_answer.asp?id=013>旅行选店论情人</a></td>
                      <td width="33%" height="17"><a href=z_answer.asp?id=014>你能驾驭爱情吗?</a></td>
                      <td width="34%" height="17"><a href=z_answer.asp?id=015>他是否值得你托付终身</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="17"><a href=z_answer.asp?id=016>你是性压抑者吗</a></td>
                      <td width="33%" height="17"><a href=z_answer.asp?id=017>你是哪种类型的丈夫</a></td>
                      <td width="34%" height="17"><a href=z_answer.asp?id=018>恋爱你会付出多少</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="17"><a href=z_answer.asp?id=019>你的恋人好色与否</a></td>
                      <td width="33%" height="17"><a href=z_answer.asp?id=020>十张纸要你丢掉时</a></td>
                      <td width="34%" height="17"><a href=z_answer.asp?id=021>你了解她么</td>
                    </tr>
                    <tr>
                      <td width="33%" height="17"><a href=z_answer.asp?id=022>对帅哥有那些联想呢</a></td>
                      <td width="33%" height="17"><a href=z_answer.asp?id=023>你们挑的是那一种戒指？</a></td>
                      <td width="34%" height="17"><a href=z_answer.asp?id=024>婚姻危机</a></td>
                    </tr>
                  </table>
                  </center>
                </div>
                </td>
              </tr>
              <tr>
                <td width="80" height="215" valign="top" align="center">
                <p style="margin-top: 5">
                <img border="0" src="ontest/zonghe.GIF"></p>
                <p><font color="#0000FF">[</font> <font color="#FF0000">综合测试</font>
                <font color="#0000FF">]</font><p>　</td>
                <td width="500" height="215" valign="top">
                <div align="center">
                  <center>
                  <table border="0" cellspacing="3" width="95%" id="AutoNumber7" cellpadding="0" height="47" background="ontest/bg3.gif">
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=101>你是个乐观的人吗？</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=102>信心测试</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=103>你有变革意识吗？</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=104>你适合的工作</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=105>是否乐于从事与人打交道</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=106>你胸怀大志吗？</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=107>你是个求安稳的人吗？</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=108>测试你的沟通能力</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=109>善于化解与上司的矛盾吗？</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=110>测验你的家庭是否美满</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=111>为啥最终能赢是他而不是你</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=112>测试投资人的&quot;理财EQ&quot;</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=113>你记得...吗？</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=114>选以下哪个树枝？</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=115>你将来的生财计划</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=116>测验你的家庭是否美满？</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=117>意志力自测</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=118>你是个爱慕虚荣的人吗？</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=119>你是好强固执的人吗？</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=120>心理适应性测试</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=121>责任感测试</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=122>由颜色判别个性爱好</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=123>性向测试</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=124>自主性测试</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=125>你是轻松兴奋的人吗？</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=126>意志力测试</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=127>你有焦虑情绪吗？</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=128>你特别敏感吗？</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=129>你点菜时通常是？</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=130>拿着什么东西而逃呢？</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=131>你是什么人?</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=132>习惯可以知道一个人</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=133>喜欢什么人坐在你的身旁？</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=134>你有什么吸引力？</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=135>你是胸怀大志的人吗？</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=136>你是轻松兴奋的人吗？</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=137>你特别敏感吗？</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=138>你的成功指数有多少？</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=139>测试你的沟通能力</a></td>
                    </tr>
                    </table>
                  </center>
                </div>
                </td>
              </tr>
              <tr>
                <td width="80" height="189" valign="top" align="center">
                <p style="margin-top: 5">
                <img border="0" src="ontest/gexing.GIF"></p>
                <p><font color="#0000FF">[</font> <font color="#FF0000">个性测试</font>
                <font color="#0000FF">]</font><p>　</td>
                <td width="500" height="189" valign="top">
                <div align="center">
                  <center>
                  <table border="0" cellspacing="3" width="95%" id="AutoNumber7" cellpadding="0" height="17" background="ontest/bg3.gif">
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=201>你留给人的第一印象如何？</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=202>你与朋友们处得愉快吗？</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=203>你很精明事故吗？</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=204>男性魅力测验</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=205>你是受欢迎的人吗？</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=206>处理公共关系水平测验</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=207>羞却测试</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=208>你是否心平气和？</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=209>女性魅力测验？</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=210>魅力是一种综合测试</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=211>社交综合</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=212>插几根蜡烛？</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=213>男与女凹凸配对</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=214>你的人生主次</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=215>你认为这是何种建筑物？</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=216>女孩到底在说什么？</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=217>把钱摇出来</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=218>生熟鸡蛋的问题</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=219>鬼火扮靓</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=220>新机场一日游</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=221>做夜鬼刨书</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=222>我是谁？</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=223>做个大富家</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=224>哪个类型的男孩最吸引你</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=225>天生爱情狂</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=226>爱上我吗？</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=227>人际关系的测试</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=228>哪个类型的女孩最吸引你</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=229>你的欲求度如何</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=230>你是男子汉吗</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=231>你处理问题的能力强吗</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=232>你属于那一类型妻子</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=233>评估你的性魅力与爱的能力</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=234>你对你的性生活感到满足吗</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=235>你与性伴是否有很好的沟通</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=236>他是怎样的人呢？</a></td>
                    </tr>
                  </table>
                  </center>
                </div>
                </td>
              </tr>
            </table>
            </center>
          </div>
          </td>
        </tr>
      </table>
      </center>
    </div>
    </td>
  </tr>
</table>
</body>

</html>
<%
response.write "</CENTER></td></tr></table>"
end if
call footer()
%>