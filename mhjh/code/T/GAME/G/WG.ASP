<%
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../../error.asp?id=016"
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
conn.Execute ("update 用户 set 银两=银两-100 where 姓名='"&username&"'")
conn.Close
set conn=nothing
%>
<HTML><HEAD><TITLE>文 王 神 卦</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb_2312-80">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<STYLE> {
	FONT-SIZE: 9pt
}
 {
	TEXT-DECORATION: none
}
</STYLE>
</HEAD>
<BODY text=#000000 vLink=#ff0000 aLink=#00ff00 link=#ff0000 bgColor=#000000 leftMargin=0 topMargin=10 oncontextmenu=self.event.returnValue=false>
<SCRIPT language=JScript.Encode src="文 王 神 卦.files/1sg.js"></SCRIPT>
<FORM name=FrontPage_Form1 onsubmit="return FrontPage_Form1_Validator(this)">
<DIV align=center>
<CENTER>
<TABLE style="FONT-SIZE: 9pt" cellSpacing=0 borderColorDark=#ffffff 
cellPadding=0 width=450 bgColor=#8cc6ff borderColorLight=#000000 border=1 height="247">
  <TBODY>
  <TR>
    <TD align=middle width=876 height=89 colspan="2"><BR>
      <CENTER>
      <TABLE style="FONT-SIZE: 9pt" height=20 cellSpacing=0 cellPadding=0 
      width=686 bgColor=#808080 border=0>
        <TBODY>
        <TR>
          <TD width=686><FONT color=#ffffff>&nbsp;&nbsp;您的位置:《 文 王 神 卦          
            》</FONT></TD></TR></TBODY></TABLE></CENTER>         
      <DIV align=left><BR>&nbsp;&nbsp; 请选择你要问的项目： </DIV>　</TD>       
  </TR>        
  <TR align=middle>        
    <TD align=middle width=638 height="116">        
      <DIV align=center>        
      <CENTER>        
      <TABLE style="FONT-SIZE: 9pt" cellSpacing=0 cellPadding=0 width=538         
      border=0>        
        <TBODY>        
        <TR>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,11) type=radio value=V65         
            name=R1>晴雨-11</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,12) type=radio value=V66         
            name=R1>升学-12</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,13) type=radio value=V67         
            name=R1>求学-13</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67>        
            <DIV align=center>        
            <CENTER>        
            <P><SPAN style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,24) type=radio value=V129         
            name=R1>取讨-14</FONT></SPAN></P></CENTER></DIV></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,15) type=radio value=V69         
            name=R1>招婿-15</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,16) type=radio value=V70         
            name=R1>命运-16</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=68><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,17) type=radio value=V71         
            name=R1>请医-17</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=68><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,18) type=radio value=V72         
            name=R1>考试-18</FONT></SPAN></TD></TR>        
        <TR align=middle>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,21) type=radio value=V73         
            name=R1>移居-21</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,22) type=radio value=V74         
            name=R1>会事-22</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,23) type=radio value=V75         
            name=R1>谋事-23</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,24) type=radio value=V76         
            name=R1>父母-24</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,25) type=radio value=V77         
            name=R1>房屋-25</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,26) type=radio value=V78         
            name=R1>分家-26</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=68><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,27) type=radio value=V79         
            name=R1>疾病-27</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=68><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,28) type=radio value=V80         
            name=R1>天花-28</FONT></SPAN></TD></TR>        
        <TR align=middle>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,31) type=radio value=V88         
            name=R1>求财-31</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,32) type=radio value=V87         
            name=R1>借财-32</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,33) type=radio value=V86         
            name=R1>购畜-33</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,34) type=radio value=V85         
            name=R1>开店-34</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,35) type=radio value=V84         
            name=R1>回乡-35</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,36) type=radio value=V83         
            name=R1>放帐-36</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=68><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,37) type=radio value=V82         
            name=R1>坟墓-37</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=68><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,38) type=radio value=V81         
            name=R1>赌钱-38</FONT></SPAN></TD></TR>        
        <TR align=middle>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,41) type=radio value=V96         
            name=R1>秋收-41</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,42) type=radio value=V95         
            name=R1>寻馆-42</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,43) type=radio value=V94         
            name=R1>求子-43</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,44) type=radio value=V93         
            name=R1>出行-44</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,45) type=radio value=V92         
            name=R1>取鱼-45</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,46) type=radio value=V91         
            name=R1>夜梦-46</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=68><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,47) type=radio value=V90         
            name=R1>入赘-47</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=68><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,48) type=radio value=V89         
            name=R1>口舌-48</FONT></SPAN></TD></TR>        
        <TR align=middle>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,51) type=radio value=V104         
            name=R1>脱货-51</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,52) type=radio value=V103         
            name=R1>见官-52</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,53) type=radio value=V102         
            name=R1>春蚕-53</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,54) type=radio value=V101         
            name=R1>解粮-54</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,55) type=radio value=V100         
            name=R1>起造-55</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,56) type=radio value=V99         
            name=R1>生意-56</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=68><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,57) type=radio value=V98         
            name=R1>文书-57</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=68><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,58) type=radio value=V97         
            name=R1>诉讼-58</FONT></SPAN></TD></TR>        
        <TR align=middle>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,61) type=radio value=V112         
            name=R1>升迁-61</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,62) type=radio value=V111         
            name=R1>寻人-62</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,63) type=radio value=V110         
            name=R1>纳监-63</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,64) type=radio value=V109         
            name=R1>和事-64</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,65) type=radio CHECKED value=V108         
            name=R1>婚姻-65</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,66) type=radio value=V107         
            name=R1>怀孕-66</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=68><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,67) type=radio value=V106         
            name=R1>交易-67</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=68><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,68) type=radio value=V105         
            name=R1>取妾-68</FONT></SPAN></TD></TR>        
        <TR align=middle>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,71) type=radio value=V120         
            name=R1>田产-71</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,72) type=radio value=V119         
            name=R1>家信-72</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,73) type=radio value=V118         
            name=R1>买官-73</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,74) type=radio value=V117         
            name=R1>告状-74</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,75) type=radio value=V116         
            name=R1>买卖-75</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,76) type=radio value=V115         
            name=R1>求官-76</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=68><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,77) type=radio value=V114         
            name=R1>跟官-77</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=68><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,78) type=radio value=V113         
            name=R1>讨仆-78</FONT></SPAN></TD></TR>        
        <TR align=middle>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,81) type=radio value=V128         
            name=R1>寿命-81</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,82) type=radio value=V127         
            name=R1>家宅-82</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,83) type=radio value=V126         
            name=R1>合伙-83</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,84) type=radio value=V125         
            name=R1>失物-84</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,85) type=radio value=V124         
            name=R1>走失-85</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=67><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,86) type=radio value=V123         
            name=R1>行人-86</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=68><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,87) type=radio value=V122         
            name=R1>手艺-87</FONT></SPAN></TD>        
          <TD vAlign=center align=middle width=68><SPAN         
            style="LINE-HEIGHT: 4px"><FONT color=#000000><INPUT         
            onclick=setwz(this,88) type=radio value=V121         
            name=R1>解人-88</FONT></SPAN></TD></TR></TBODY></TABLE></CENTER></DIV></TD>        
    <TD align=middle width=638 height="116">        
      　<img border="0" src="r01.jpg" width="150"></TD></TR>  
  <TR align=middle>  
    <TD align=middle width=638 bgColor=#808080 height=36><FONT   
      color=#ffffff><INPUT type=hidden name=T3> 随意输入一个数字（0-999）</FONT><INPUT   
      style="BORDER-RIGHT: rgb(192,192,192) 1px solid; BORDER-TOP: rgb(192,192,192) 1px solid; FONT-SIZE: 9pt; BORDER-LEFT: rgb(192,192,192) 1px solid; BORDER-BOTTOM: rgb(192,192,192) 1px solid"   
      onfocus="this.form.T1.value=''" size=12 value=请输入数字 name=T1><FONT   
      color=#000000> <INPUT style="BORDER-RIGHT: rgb(255,255,255) 1px dotted; BORDER-TOP: rgb(255,255,255) 1px dotted; FONT-SIZE: 9pt; BORDER-LEFT: rgb(255,255,255) 1px dotted; COLOR: rgb(255,255,255); BORDER-BOTTOM: rgb(255,255,255) 1px dotted; BACKGROUND-COLOR: rgb(192,192,192)" onclick=doform(this.form) type=button value=" 确  定"></FONT>  
      <SCRIPT language=javascript>         
</SCRIPT>         
       </TD>         
    <TD align=middle width=638 bgColor=#808080 height="36">　         
       </TD></TR></TBODY></TABLE></CENTER></DIV>         
<DIV align=center style="width: 764; height: 1"><INPUT type=hidden name=T2>          
</DIV></FORM>        
<p align="center"><font color="#FFFFFF">你花费了100两银子来到&lt;</font><FONT color=#ffffff>文 王 神 卦&gt;</FONT></p>    
</BODY></HTML>       
