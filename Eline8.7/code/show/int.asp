<%
dim LocalPic,PicPath
LocalPic=0
if LocalPic=1 then
	PicPath="userface/"
else
	PicPath="http://qqshow-item.tencent.com/"
end if
sub DispPageNum(CurPage, PageCount, URLPrefix, URLPostfix)
 dim p,ii
 if PageCount=0 then
  response.write "无 "
 else
  p=(CurPage-1) \ 10
  if CurPage=1 then
   response.write "<SPAN class=pagenumstatic><font face=webdings color=red>9</font></SPAN>  "
  else
   response.write "<SPAN class=pagenum><a href="&URLPrefix&"1"&URLPostfix&" title=首页><font face=webdings>9</font></a></SPAN> "
  end if
  if p>0 then response.write "<SPAN class=pagenum><a href="&URLPrefix&Cstr(p*10)&URLPostfix&" title=上十页><font face=webdings>7</font></a></SPAN> "
  response.write "<b>"
  for ii=p*10+1 to P*10+10
   if ii=CurPage then
     response.write "<SPAN class=pagenumstatic><B><font color=red>"+Cstr(ii)+"</font></B></SPAN> "
   else
    response.write "<SPAN class=pagenum><a href="&URLPrefix&Cstr(ii)&URLPostfix&">"+Cstr(ii)+"</a></SPAN> "
   end if
   if ii=PageCount then exit for
  next
  if ii>p*10+10 then ii=ii-1
  response.write "</b>"
  if ii<PageCount then response.write "<SPAN class=pagenum><a href="&URLPrefix&Cstr(ii+1)&URLPostfix&" title=下十页><font face=webdings>8</font></a></SPAN> "
  if CurPage=PageCount then
   response.write "<SPAN class=pagenumstatic><font face=webdings color=red>:</font></SPAN>"
  else
   response.write "<SPAN class=pagenum><a href="&URLPrefix&Cstr(PageCount)&URLPostfix&" title=尾页><font face=webdings>:</font></a></SPAN> "
  end if
 end if
end sub

function GetVisualStr(CurVisualStr,CurVisualSex,CurVisualPeriod)
 dim VisualStr,VisualSex,VisualPeriod
 dim VisualSplit,PeriodSplit
 dim i,TempSplit
 VisualStr=CurVisualStr
 VisualSex=CurVisualSex
 VisualPeriod=CurVisualPeriod
 if isnull(VisualStr) or VisualStr="" then
  if VisualSex=1 then
   VisualStr="||||||14.7|13.8|12.9||11.11||10.13|9.14||||8.18|||||||"
  else
   VisualStr="||||||7.7|6.8|5.9||4.11||3.13|2.14||||1.18|||||||"
  end if
 else
  if ubound(split(VisualStr,"|"))<>24 then
   if VisualSex=1 then
    VisualStr="||||||14.7|13.8|12.9||11.11||10.13|9.14||||8.18|||||||"
   else
    VisualStr="||||||7.7|6.8|5.9||4.11||3.13|2.14||||1.18|||||||"
   end if
  else
   if not isnull(VisualPeriod) then
    PeriodSplit=split(VisualPeriod,"|")
    if ubound(PeriodSplit)=24 then
     VisualSplit=split(VisualStr,"|")
     VisualStr=""
     for i=0 to 24
      if VisualSplit(i)<>"" then
       TempSplit=split(VisualSplit(i),".")
       if cint(TempSplit(0))>20 then
        if PeriodSplit(i)<>"" then
         TempSplit=split(PeriodSplit(i),",")
         if ubound(TempSplit)=1 then
          if len(TempSplit(0))=6 and isnumeric(TempSplit(0)) and isnumeric(TempSplit(1)) then
           if datediff("d",CDate("20"&left(TempSplit(0),2)&"-"&mid(TempSplit(0),3,2)&"-"&right(TempSplit(0),2)),now())>cint(TempSplit(1)) and cint(TempSplit(1))>0 then
            select case i
            case 6
             if VisualSex=1 then
              VisualSplit(i)="14.7"
             else
              VisualSplit(i)="7.7"
             end if
            case 7
             if VisualSex=1 then
              VisualSplit(i)="13.8"
             else
              VisualSplit(i)="6.8"
             end if
            case 8
             if VisualSex=1 then
              VisualSplit(i)="12.9"
             else
              VisualSplit(i)="5.9"
             end if
            case 10
             if VisualSex=1 then
              VisualSplit(i)="11.11"
             else
              VisualSplit(i)="4.11"
             end if
            case 12
             if VisualSex=1 then
              VisualSplit(i)="10.13"
             else
              VisualSplit(i)="3.13"
             end if
            case 13
             if VisualSex=1 then
              VisualSplit(i)="9.14"
             else
              VisualSplit(i)="2.14"
             end if
            case 17
             if VisualSex=1 then
              VisualSplit(i)="8.18"
             else
              VisualSplit(i)="1.18"
             end if
            case else
             VisualSplit(i)=""
            end select
           end if
          end if
         end if
        end if
       end if
      end if
      if i>0 then VisualStr=VisualStr&"|"
      VisualStr=VisualStr&VisualSplit(i)
     next
    end if
   end if
  end if
 end if
 GetVisualStr=VisualStr
end function

function checkStr(str)
	if isnull(str) then
		checkStr = ""
		exit function 
	end if
	checkStr=replace(str,"'","''")
end function

function GetLayerStr(CurVisualSplit)
 dim i,LayerStr
 LayerStr=""
 for i=0 to ubound(CurVisualSplit)
  if not isnull(CurVisualSplit(i)) and trim(CurVisualSplit(i))<>"" then
   if LayerStr<>"" then LayerStr=LayerStr&"_"
   LayerStr=LayerStr&(i+1)
  end if
 next
 GetLayerStr=LayerStr
end function

function isCanShowVisual(CurVisualSplit)
 dim TempSplit
 if CurVisualSplit(6)="" or CurVisualSplit(7)="" or CurVisualSplit(8)="" or CurVisualSplit(10)="" or CurVisualSplit(12)="" or CurVisualSplit(13)="" or CurVisualSplit(17)="" then
  isCanShowVisual=False
  exit function
 end if
 TempSplit=split(CurVisualSplit(7),".")
 if TempSplit(0)="6" or TempSplit(0)="13" then
  isCanShowVisual=False
  exit function
 end if
 TempSplit=split(CurVisualSplit(8),".")
 if TempSplit(0)="5" or TempSplit(0)="12" then
  isCanShowVisual=False
  exit function
 end if
 isCanShowVisual=True
end function%>