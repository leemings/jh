<%
function astro(birth)
	on error resume next
	dim birthday,birthmonth,birthyear
	birthday=day(birth)
	birthmonth=month(birth)
	select case birthmonth
	case 1
		if birthday>=21 then
			astro="<img src=star/z11.gif alt=Ë®Æ¿×ù"&birth&">"
		else
			astro="<img src=star/z10.gif alt=Ä§ôÉ×ù"&birth&">"
		end if
	case 2
		if birthday>=20 then
			astro="<img src=star/z12.gif alt=Ë«Óã×ù"&birth&">"
		else
			astro="<img src=star/z11.gif alt=Ë®Æ¿×ù"&birth&">"
		end if
	case 3
		if birthday>=21 then
			astro="<img src=star/z1.gif alt=°×Ñò×ù"&birth&">"
		else
			astro="<img src=star/z12.gif alt=Ë«Óã×ù"&birth&">"
		end if
	case 4
		if birthday>=21 then
			astro="<img src=star/z2.gif alt=½ğÅ£×ù"&birth&">"
		else
			astro="<img src=star/z1.gif alt=°×Ñò×ù"&birth&">"
		end if
	case 5
		if birthday>=22 then
			astro="<img src=star/z3.gif alt=Ë«×Ó×ù"&birth&">"
		else
			astro="<img src=star/z2.gif alt=½ğÅ£×ù"&birth&">"
		end if
	case 6
		if birthday>=22 then
			astro="<img src=star/z4.gif alt=¾ŞĞ·×ù"&birth&">"
		else
			astro="<img src=star/z3.gif alt=Ë«×Ó×ù"&birth&">"
		end if
	case 7
		if birthday>=23 then
			astro="<img src=star/z5.gif alt=Ê¨×Ó×ù"&birth&">"
		else
			astro="<img src=star/z4.gif alt=¾ŞĞ·×ù"&birth&">"
		end if
	case 8
		if birthday>=24 then
			astro="<img src=star/z6.gif alt=´¦Å®×ù"&birth&">"
		else
			astro="<img src=star/z5.gif alt=Ê¨×Ó×ù"&birth&">"
		end if
	case 9
		if birthday>=24 then
			astro="<img src=star/z7.gif alt=Ìì³Ó×ù"&birth&">"
		else
			astro="<img src=star/z6.gif alt=´¦Å®×ù"&birth&">"
		end if
	case 10
		if birthday>=24 then
			astro="<img src=star/z8.gif alt=ÌìĞ«×ù"&birth&">"
		else
			astro="<img src=star/z7.gif alt=Ìì³Ó×ù"&birth&">"
		end if
	case 11
		if birthday>=23 then
			astro="<img src=star/z9.gif alt=ÉäÊÖ×ù"&birth&">"
		else
			astro="<img src=star/z8.gif alt=ÌìĞ«×ù"&birth&">"
		end if
	case 12
		if birthday>=22 then
			astro="<img src=star/z10.gif alt=Ä§ôÉ×ù"&birth&">"
		else
			astro="<img src=star/z9.gif alt=ÉäÊÖ×ù"&birth&">"
		end if
	case else
		astro=""
	end select
	birthyear=(year(birth)-1900) mod 12
	select case birthyear
	case 0
		astro="<img src=images/img_shengxiao/"&birthyear+1&".gif alt=Êó> "&astro
	case 1
		astro="<img src=images/img_shengxiao/"&birthyear+1&".gif alt=Å£> "&astro
	case 2
		astro="<img src=images/img_shengxiao/"&birthyear+1&".gif alt=»¢> "&astro
	case 3
		astro="<img src=images/img_shengxiao/"&birthyear+1&".gif alt=ÍÃ> "&astro
	case 4
		astro="<img src=images/img_shengxiao/"&birthyear+1&".gif alt=Áú> "&astro
	case 5
		astro="<img src=images/img_shengxiao/"&birthyear+1&".gif alt=Éß> "&astro
	case 6
		astro="<img src=images/img_shengxiao/"&birthyear+1&".gif alt=Âí> "&astro
	case 7
		astro="<img src=images/img_shengxiao/"&birthyear+1&".gif alt=Ñò> "&astro
	case 8
		astro="<img src=images/img_shengxiao/"&birthyear+1&".gif alt=ºï> "&astro
	case 9
		astro="<img src=images/img_shengxiao/"&birthyear+1&".gif alt=¼¦> "&astro
	case 10
		astro="<img src=images/img_shengxiao/"&birthyear+1&".gif alt=¹·> "&astro
	case 11
		astro="<img src=images/img_shengxiao/"&birthyear+1&".gif alt=Öí> "&astro
	case else
	end select
end function
%>