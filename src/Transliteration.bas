Attribute VB_Name = "Transliteration"
Option Explicit
Option Compare Binary

'В основном правила транслитерации схемы Мосметро: https://dangry.ru/iuliia/mosmetro/
'Исключения: буква Х и украинские буквы.
Function Translit(Src As String) As String
    Dim I As Integer
    Dim Dst() As String
    Dim LenSrc As Integer
    Dim Letter As String
    Dim PrevLetter As String
    Dim NextLetter As String
    
    If Src = "" Then Exit Function
    LenSrc = Len(Src)
    ReDim Dst(LenSrc - 1)
    For I = 0 To LenSrc - 1
        If I = 0 Then
            PrevLetter = " "
        Else
            PrevLetter = Mid(Src, I, 1)
        End If
        Letter = Mid(Src, I + 1, 1)
        NextLetter = Mid(Src, I + 2, 1)
        Select Case Letter
      
            Case "А"
                Dst(I) = "A"
            Case "а"
                Dst(I) = "a"
              
            Case "Б"
                Dst(I) = "B"
            Case "б"
                Dst(I) = "b"
              
            Case "В"
                Dst(I) = "V"
            Case "в"
                Dst(I) = "v"
              
            Case "Г"
                Dst(I) = "G"
            Case "г"
                Dst(I) = "g"
              
            Case "Д"
                Dst(I) = "D"
            Case "д"
                Dst(I) = "d"
              
            Case "Е"
                Dst(I) = "E"
            Case "е"
                Dst(I) = "e"
              
            Case "Ё"
                Select Case PrevLetter
                    Case "Ь", "ь", "Ъ", "ъ", " "
                        Dst(I) = "Yo"
                    Case Else
                        Dst(I) = "E"
                End Select
            Case "ё"
                Select Case PrevLetter
                    Case "Ь", "ь", "Ъ", "ъ", " "
                        Dst(I) = "yo"
                    Case Else
                        Dst(I) = "e"
                End Select
              
            Case "Ж"
                Dst(I) = "Zh"
            Case "ж"
                Dst(I) = "zh"
              
            Case "З"
                Dst(I) = "Z"
            Case "з"
                Dst(I) = "z"
              
            Case "И"
                Dst(I) = "I"
            Case "и"
                Dst(I) = "i"
              
            Case "Й"
                Select Case PrevLetter
                    Case "Ы", "ы", "И", "и"
                        Dst(I - 1) = "Y"
                        Dst(I) = ""
                    Case Else
                        Dst(I) = "Y"
                End Select
            Case "й"
                Select Case PrevLetter
                    Case "Ы", "ы", "И", "и"
                        Dst(I - 1) = "y"
                        Dst(I) = ""
                    Case Else
                        Dst(I) = "y"
                End Select
              
            Case "К"
                Dst(I) = "K"
            Case "к"
                Dst(I) = "k"
              
            Case "Л"
                Dst(I) = "L"
            Case "л"
                Dst(I) = "l"
              
            Case "М"
                Dst(I) = "M"
            Case "м"
                Dst(I) = "m"
              
            Case "Н"
                Dst(I) = "N"
            Case "н"
                Dst(I) = "n"
              
            Case "О"
                Dst(I) = "O"
            Case "о"
                Dst(I) = "o"
              
            Case "П"
                Dst(I) = "P"
            Case "п"
                Dst(I) = "p"
              
            Case "Р"
                Dst(I) = "R"
            Case "р"
                Dst(I) = "r"
              
            Case "С"
                Dst(I) = "S"
            Case "с"
                Dst(I) = "s"
              
            Case "Т"
                Dst(I) = "T"
            Case "т"
                Dst(I) = "t"
              
            Case "У"
                Dst(I) = "U"
            Case "у"
                Dst(I) = "u"
              
            Case "Ф"
                Dst(I) = "F"
            Case "ф"
                Dst(I) = "f"
              
            Case "Х"
                Dst(I) = "X"  'Kh
            Case "х"
                Dst(I) = "x"  'kh
              
            Case "Ц"
                Select Case PrevLetter
                    Case "Т", "т"
                        Dst(I) = "S"
                    Case Else
                        Dst(I) = "Ts"
                End Select
            Case "ц"
                Select Case PrevLetter
                    Case "Т", "т"
                        Dst(I) = "s"
                    Case Else
                        Dst(I) = "ts"
              End Select
              
            Case "Ч"
                Dst(I) = "Ch"
            Case "ч"
                Dst(I) = "ch"
              
            Case "Ш"
                Dst(I) = "Sh"
            Case "ш"
                Dst(I) = "sh"
              
            Case "Щ"
                Dst(I) = "Sch"
            Case "щ"
                Dst(I) = "sch"
              
            Case "Ь"
                Select Case NextLetter
                    Case "А", "а", "Е", "е", "И", "и", "О", "о", "У", "у", "Э", "э"
                        Dst(I) = "Y"
                    Case Else
                        Dst(I) = ""
                End Select
            Case "ь"
                Select Case NextLetter
                    Case "А", "а", "Е", "е", "И", "и", "О", "о", "У", "у", "Э", "э"
                        Dst(I) = "y"
                    Case Else
                        Dst(I) = ""
                End Select
              
            Case "Ы"
                Dst(I) = "Y"
            Case "ы"
                Dst(I) = "y"
              
            Case "Ъ"
                Select Case NextLetter
                    Case "А", "а", "Е", "е", "И", "и", "О", "о", "У", "у", "Э", "э"
                        Dst(I) = "Y"
                    Case Else
                        Dst(I) = ""
                End Select
            Case "ъ"
                Select Case NextLetter
                    Case "А", "а", "Е", "е", "И", "и", "О", "о", "У", "у", "Э", "э"
                        Dst(I) = "y"
                    Case Else
                        Dst(I) = ""
                End Select
              
            Case "Э"
                Dst(I) = "E"
            Case "э"
                Dst(I) = "e"
              
            Case "Ю"
                Dst(I) = "Yu"
            Case "ю"
                Dst(I) = "yu"
              
            Case "Я"
                Dst(I) = "Ya"
            Case "я"
                Dst(I) = "ya"
              
            'Украинские буквы
            Case "Ґ"
                Dst(I) = "G"
            Case "ґ"
                Dst(I) = "g"
              
            Case "Є"
                Dst(I) = "E"
            Case "є"
                Dst(I) = "e"
              
            Case "Ї"
                Select Case PrevLetter
                    Case " "
                        Dst(I) = "Yi"
                    Case Else
                        Dst(I) = "I"
                End Select
            Case "ї"
                Select Case PrevLetter
                    Case " "
                        Dst(I) = "yi"
                    Case Else
                        Dst(I) = "i"
                End Select
              
            Case "І"
                Dst(I) = "I"
            Case "і"
                Dst(I) = "i"
            
            Case Else
                Dst(I) = Letter
        End Select
    Next
    
    Translit = Join(Dst, "")
End Function
