<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
%>
<html><head><title>�ռ�����</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head><LINK href="images/diary.css" rel=stylesheet type=text/css>
<body bgcolor="#FFFFFF" text="#000000">
<DIV ID="overDiv" STYLE="position:absolute; visibility:hide; z-index: 1;"></DIV>
<SCRIPT LANGUAGE="JavaScript" src="images/overlib.js"></SCRIPT>
<table align=center border=0 cellpadding=0 cellspacing=0 width=574>
<tr><td><img height=13 src="images/top.gif" width=574></td></tr>
<tr><td align=right background=images/top1.gif><table border=0 cellpadding=0 cellspacing=0 width="94%">
<tr><td align=right><table border=0 cellpadding=0 cellspacing=0 width="80%">
<tr align=right><td><a href="mydiary.asp"><img alt=�����ҵ��ռǱ� border=0 height=21 src="images/anniu1.gif" width=81></a></td>
<td><a href="list.asp"><img alt=�����ռ� border=0 height=21 src="images/anniu2.gif" width=80></a></td>
<td><a href="top.asp?fs=1"><img border=0 height=21 src="images/anniu3.gif" width=91></a></td>
<td><a href="top.asp?fs=2"><img alt=�ռ����� border=0 height=21 src="images/anniu4.gif" width=78></a></td></tr>
</table></td><td rowspan=3 valign=bottom width=26><img height=289 src="images/right.gif" width=8></td>
</tr><%If Request.QueryString("CurPage") = "" or Request.QueryString("CurPage") = 0 then
	CurPage = 1
Else
	CurPage = CINT(Request.QueryString("CurPage"))
End If
sub diaryclose(mess)
	rs1.close
	set rs1=nothing
	conn1.close
	set conn1=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��"& mess &"');location.href = 'list.asp';</script>"
	response.end
end sub
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
dim mylb(3)
rsfs=Request.QueryString("fs")
select case rsfs
case "1"
	rs1.Open "Select * From �ռ��û� Order By �Ĳ� DESC", conn1, 1,1
	if rs1.eof and rs1.bof then
		call diaryclose("���û����������⣬�޷��鿴��")
	end if
'	bb1=rs1("�ռǱ���")
'	bb4=rs1("�ʻ�")
'	bb5=rs1("����")
'	bb6=rs1("�ռ���")
'	bb7=rs1("������")
'	bb8=rs1("��������")
'	bb9=rs1("�Ĳ�")
'	mylb(1)=rs1("lb1")
'	mylb(2)=rs1("lb2")
'	mylb(3)=rs1("lb3")
case "2"
	rs1.Open "Select * From �ռ� Order By (�ʻ�-����) DESC", conn1, 1,1
	if rs1.eof and rs1.bof then
		call diaryclose("���û����������⣬�޷��鿴��")
	end if
end select
if rs1.eof and rs1.bof then
%><tr><td align=center >�Բ���û���ҵ��κμ�¼!</td></tr><%
else
rs1.PageSize=15'����ÿҳ��¼��,���޸�          
Dim TotalPages              
TotalPages = rs1.PageCount              
If CurPage>rs1.Pagecount Then               
	CurPage=rs1.Pagecount              
end if                            
rs1.AbsolutePage=CurPage
rs1.CacheSize = rs1.PageSize'��������¼��  
Dim Totalcount
Totalcount =INT(rs1.recordcount)              
StartPageNum=1              
do while StartPageNum+10<=CurPage              
StartPageNum=StartPageNum+10              
Loop              
EndPageNum=StartPageNum+9              
If EndPageNum>rs1.Pagecount then EndPageNum=rs1.Pagecount
%><tr valign="top"><td>
 <table align=center bgcolor=#e0d2c5 border=3 cellpadding=2 cellspacing=0 width="98%" >
<%I=0
p=rs1.PageSize*(Curpage-1)              
do while (Not rs1.Eof) and (I<rs1.PageSize)              
p=p+1
if rsfs="2" then
	select case rs1("����")
	case "1"
		bm="��������"
	case "2"
		bm="��������!"
	case "3"
		bm="ƾ����鿴"
	case "4"
		bm="["&rs1("��������") &"]�鿴"
	case "5"
		bm="["&rs1("��������")&"]������"
	case "6"
		bm="["&rs1("��������")&"]�鿴"
	end select
	if i/2=int(i/2) then%>
              <tr align=left bgcolor='cccccc' onmouseover=this.bgColor='#00B0DD'; onmouseout=this.bgColor='cccccc'> 
                <td width="63%" > 
                  <img height=14 src="images/hot.gif"width=11><a href="show.asp?id=<%=rs1("id")%>&wj=list.asp" onMouseOver="drs('<div align=center><b>����Ϣ���ϡ�</b></div><img src=images/xianhua.gif width=16 height=15>:<font color=#ff0000><%=rs1("�ʻ�")%></font>��<br><img src=images/jidan.gif width=16 height=15>:<font color=#ff0000><%=rs1("����")%></font>��<br>����:<font color=#ff0000><%=rs1("������")%></font>��<br>��ʽ:<%=bm%>');return true;" onMouseOut="nd(); return true;"><%=rs1("�ռǱ���")%></a>
                </td>
                <td width="37%"> 
                  <a href="list.asp?fs=1&name=<%=rs1("�û���")%>" target="_blank"><img src="images/gd.gif" width="13" height="13" border="0" alt="�鿴����ר����"></a>&nbsp;����:<a href="mydiary.asp?name=<%=rs1("�û���")%>" title="�鿴�����ռǱ�" target="_blank"><%=rs1("�û���")%></a>&nbsp;<%=rs1("��д����")%>
                </td>
              </tr>
              <%else%>
              <tr align=left bgcolor='efefef' onmouseover=this.bgColor='#00B0DD'; onmouseout=this.bgColor='efefef'> 
                <td width="63%" > 
                  <img height=14 src="images/hot.gif"width=11><a href="show.asp?id=<%=rs1("id")%>&wj=list.asp" onMouseOver="drs('<div align=center><b>����Ϣ���ϡ�</b></div>�ʻ�:<font color=#ff0000><%=rs1("�ʻ�")%></font>��<br>����:<font color=#ff0000><%=rs1("����")%></font>��<br>����:<font color=#ff0000><%=rs1("������")%></font>��<br>��ʽ:<%=bm%>');return true;" onMouseOut="nd(); return true;"><%=rs1("�ռǱ���")%></a>
                </td>
                <td width="37%"> 
                  <a href="list.asp?fs=1&name=<%=rs1("�û���")%>" target="_blank"><img src="images/gd.gif" width="13" height="13" border="0" alt="�鿴����ר����"></a>&nbsp;����:<a href="mydiary.asp?name=<%=rs1("�û���")%>" title="�鿴�����ռǱ�" target="_blank"><%=rs1("�û���")%></a>&nbsp;<%=rs1("��д����")%>
                </td>
              </tr>
              <%end if
else
	if i/2=int(i/2) then%>
              <tr align=left bgcolor='cccccc' onmouseover=this.bgColor='#00B0DD'; onmouseout=this.bgColor='cccccc'> 
                <td width="50%" ><img height=14 src="images/hot.gif"width=11>�ռǱ�����<a href="mydiary.asp?name=<%=rs1("����")%>"><%=rs1("�ռǱ���")%></a></td>
                <td width="14%">�Ĳ�:<%=rs1("�Ĳ�")%></td><td width="14%">�ռ���:<%=rs1("�ռ���")%></td>
                <td width="22%">ӵ����:<a href="mydiary.asp?name=<%=rs1("����")%>" title="�鿴�����ռǱ�" target="_blank"><%=rs1("����")%></a></td>
              </tr>
              <%else%>
              <tr align=left bgcolor='efefef' onmouseover=this.bgColor='#00B0DD'; onmouseout=this.bgColor='efefef'> 
                <td width="50%" ><img height=14 src="images/hot.gif"width=11>�ռǱ�����<a href="mydiary.asp?name=<%=rs1("����")%>"><%=rs1("�ռǱ���")%></a></td>
                <td width="14%">�Ĳ�:<%=rs1("�Ĳ�")%></td><td width="14%">�ռ���:<%=rs1("�ռ���")%></td>
                <td width="22%">ӵ����:<a href="mydiary.asp?name=<%=rs1("����")%>" title="�鿴�����ռǱ�" target="_blank"><%=rs1("����")%></a></td></tr>
              <%end if
end if
I=I+1
rs1.MoveNext
Loop
if rsfs="2" then%>
        <tr align=right bgcolor=#ede5dd> 
                  <td colSpan=2 align="right">
                    ��<%=Totalcount%>ƪ�ռ�[ 
                    <% if StartPageNum>1 then %>
                    <a href="list.asp?CurPage=<%=StartPageNum-1%>"><<</a> 
                    <%end if%>
                    <% For I=StartPageNum to EndPageNum
					if I<>CurPage then %>
                    <a href="list.asp?CurPage=<%=I%>"><%=I%></a> 
                    <% else %>
                    <b><font color=#ff0000><%=I%></font></b> 
                    <% end if %>
                    <% Next %>
                    <% if EndPageNum<rs1.Pagecount then %>
                    <a href="list.asp?CurPage=<%=EndPageNum+1%>">>></a> 
                    <%end if%>
                    ] </td>
                </tr>
<%else%>
        <tr align=right bgcolor=#ede5dd> 
                  <td colSpan=4 align="right">
                    ��<%=Totalcount%>ƪ�ռ�[ 
                    <% if StartPageNum>1 then %>
                    <a href="list.asp?CurPage=<%=StartPageNum-1%>"><<</a> 
                    <%end if%>
                    <% For I=StartPageNum to EndPageNum
					if I<>CurPage then %>
                    <a href="list.asp?CurPage=<%=I%>"><%=I%></a> 
                    <% else %>
                    <b><font color=#ff0000><%=I%></font></b> 
                    <% end if %>
                    <% Next %>
                    <% if EndPageNum<rs1.Pagecount then %>
                    <a href="list.asp?CurPage=<%=EndPageNum+1%>">>></a> 
                    <%end if%>
                    ] </td>
                </tr>
<%end if%>
            </table>
</td>
</tr>
<%end if
rs1.close
set rs1=nothing
conn1.close
set conn1=nothing%>
</table>
</td></tr><tr><td><img height=43 src="images/down.gif" width=574></td></tr></table></body></html>
