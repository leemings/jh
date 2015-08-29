 OPEN "dir.txt" FOR INPUT AS #1
 abc = 0
 OPEN "mid.asp" FOR OUTPUT AS #2
 FOR t = 1 TO 85
 LINE INPUT #1, a$
 a$ = RTRIM$(LTRIM$(a$))
 b$ = RTRIM$(LTRIM$(STR$(t)))
 aa$ = "mid(" + b$ + ")=" + CHR$(34) + a$ + CHR$(34)
 PRINT #2, aa$
 NEXT t
 CLOSE

