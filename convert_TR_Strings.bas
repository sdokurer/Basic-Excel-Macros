Function LCaseTR(ByVal text As String) As String
	Dim Letter As String
    For i = 1 To Len(text)
        Letter = Mid(text, i, 1)
		If Letter = "İ" Then
			Letter = "i"
		ElseIf Letter = "I" Then
			Letter = "ı"
		ElseIf Asc(Letter) = 138 Then
			Letter = Chr(Asc(Letter) + 16)
		ElseIf Asc(Letter) > 64 And Asc(Letter) < 91 Then
			Letter = Chr(Asc(Letter) + 32)
		ElseIf Asc(Letter) > 191 And Asc(Letter) < 215 Then
			Letter = Chr(Asc(Letter) + 32)
		ElseIf Asc(Letter) > 215 And Asc(Letter) < 223 Then
			Letter = Chr(Asc(Letter) + 32)
		Else
			Letter = Letter
		End If
		LCaseTR = LCaseTR & Letter
	Next i
End Function

Function UCaseTR(ByVal text As String) As String
	Dim Letter As String
	For i = 1 To Len(text)
		Letter = Mid(text, i, 1)
		If Letter = "i" Then
			Letter = "İ"
		ElseIf Letter = "ı" Then
			Letter = "I"
		ElseIf Asc(Letter) = 154 Then
			Letter = Chr(Asc(Letter) - 16)
		ElseIf Asc(Letter) > 96 And Asc(Letter) < 123 Then
			Letter = Chr(Asc(Letter) - 32)
		ElseIf Asc(Letter) > 223 And Asc(Letter) < 247 Then
			Letter = Chr(Asc(Letter) - 32)
		ElseIf Asc(Letter) > 247 And Asc(Letter) < 255 Then
			Letter = Chr(Asc(Letter) - 32)
		Else
			Letter = Letter
		End If
		UCaseTR = UCaseTR & Letter
	Next i
End Function

Function FCaseTR(ByVal text As String) As String
	'Dim Buyuk_Harf As String
	Dim SentenceBox() As String
	Dim FirstLetterIndex() As Integer
	Dim FirstLetter As String
	Dim z As Integer
	Dim Num As Integer

	ReDim Preserve FirstLetterIndex(0)

	z = 1

	SentenceBox = Split(text, " ")
	For i = LBound(SentenceBox) To UBound(SentenceBox)
		If SentenceBox(i) <> "" Then
			For z = 1 To Len(SentenceBox(i))
				FirstLetter = ""
				Letter = Mid(SentenceBox(i), z, 1)
				LNum = Asc(Letter)
				
			'   9-10        TAB LF
			'   13          CR
			'   32-38       SPC !"#$%&
			'   40-64       ()*+,-./0123456789:;<=>?@
			'   91-96       [\]^_`
			'   123-126     {|}~
			'   128         €
			'   130         ‚
			'   132-135     „…†‡
			'   137         ‰
			'   139         ‹
			'   145-153     ‘’“”•–—~˜™
			'   155         ›
			'   160-169     NonBreakSPC ¡¢£¤¥¦§¨©
			'   171-180     «¬­®¯°±²³´
			'   182-185     ¶·¸¹
			'   187-191     »¼½¾¿
			'   215         ×
			'   247         ÷

			'   Yukarıdaki Karakterlerden sonra eğer aşağıdaki karakterlerden biri başlarsa büyük harfe çevir

			'   65 - 90   ABCDEFGHIJKLMNOPQRSTUVWXYZ
			'   97 - 122  abcdefghijklmnopqrstuvwxyz
			'   138       Š
			'   154       š
			'   192 - 214 ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏĞÑÒÓÔÕÖ
			'   216 - 222 ØÙÚÛÜİŞ
			'   224 - 246 àáâãäåæçèéêëìíîïğñòóôõö
			'   248 - 254 øùúûüış
			'   Eğer Yukarıdaki karakterlerden biri ise ilk harf diye ata
				If LNum > 64 And LNum < 91 Or _
				LNum > 96 And LNum < 123 Or _
				LNum = 138 Or _
				LNum = 154 Or _
				LNum > 191 And LNum < 215 Or _
				LNum > 215 And LNum < 223 Or _
				LNum > 223 And LNum < 247 Or _
				LNum > 247 And LNum < 255 Then
	'   Bu kodda hata var, eğer bir kelime "/+&/()[]{}<>|@=?*-_#2%$£,;.:~´'¨ ile boşluk olmadan ayrılıyorsa ikinci kelime büyük harfle başlamıyor
					FirstLetterIndex(UBound(FirstLetterIndex)) = z
					ReDim Preserve FirstLetterIndex(UBound(FirstLetterIndex) + 1)
					
					FirstLetter = Mid(SentenceBox(i), z, 1)
					Exit For
					' Büyütülecek İlk harfi bulunca For döngüsünden çık
				End If
	'            Exit For
			Next z
			If FirstLetter = "" Then

	'       Eğer tüm harflere baktık ve yukarıdaki karakterlerden biri değilse kelimenin ilk karakterini al
				' muhakkak ilk karakter atanmalıdır, yoksa kelimeyi harf harf sıralayacağımız için sıralanmıyor.
				' z değerini daha sonra kullandığımız için "1" olarak set et
				FirstLetter = Left(SentenceBox(i), 1)
				z = 1
	'        End If
	Else
			FNum = Asc(FirstLetter)
		
	'        If FirstLetter <> "" Then
				If FirstLetter = "i" Then
					FirstLetter = "İ"
				ElseIf FirstLetter = "ı" Then
					FirstLetter = "I"
				ElseIf FNum = 154 Then
					FirstLetter = Chr(FNum - 16)
				ElseIf FNum > 96 And FNum < 123 Then
					FirstLetter = Chr(FNum - 32)
				ElseIf FNum > 223 And FNum < 247 Then
					FirstLetter = Chr(Asc(FirstLetter) - 32)
				ElseIf FNum > 247 And FNum < 255 Then
					FirstLetter = Chr(FNum - 32)
				Else
					FirstLetter = FirstLetter
				End If
				
				Mid(SentenceBox(i), z, 1) = FirstLetter
				
				If z <> Len(SentenceBox(i)) Then
					For z = z + 1 To Len(SentenceBox(i))
					
						Letter = Mid(SentenceBox(i), z, 1)
				
						If Letter = "İ" Then
							Letter = "i"
						ElseIf Letter = "I" Then
							Letter = "ı"
						ElseIf Asc(Letter) = 138 Then
							Letter = Chr(Asc(Letter) + 16)
						ElseIf Asc(Letter) > 64 And Asc(Letter) < 91 Then
							Letter = Chr(Asc(Letter) + 32)
						ElseIf Asc(Letter) > 191 And Asc(Letter) < 215 Then
							Letter = Chr(Asc(Letter) + 32)
						ElseIf Asc(Letter) > 215 And Asc(Letter) < 223 Then
							Letter = Chr(Asc(Letter) + 32)
						Else
							Letter = Letter
						End If
						
						Mid(SentenceBox(i), z, 1) = Letter
					Next z
				End If
			End If
		End If
	Next i
	For i = LBound(SentenceBox) To UBound(SentenceBox)
		If SentenceBox(i) <> "" Then
			If Sentence = "" Then
				Sentence = SentenceBox(i)
			Else
				Sentence = Sentence & " " & SentenceBox(i)
			End If
		Else
			Sentence = Sentence & " "
		End If
	Next i
	FCaseTR = Sentence
End Function
