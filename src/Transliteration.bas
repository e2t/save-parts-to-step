Attribute VB_Name = "Transliteration"
Option Explicit
Option Compare Binary

'� �������� ������� �������������� ����� ��������: https://dangry.ru/iuliia/mosmetro/
'����������: ����� � � ���������� �����.
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
      
            Case "�"
                Dst(I) = "A"
            Case "�"
                Dst(I) = "a"
              
            Case "�"
                Dst(I) = "B"
            Case "�"
                Dst(I) = "b"
              
            Case "�"
                Dst(I) = "V"
            Case "�"
                Dst(I) = "v"
              
            Case "�"
                Dst(I) = "G"
            Case "�"
                Dst(I) = "g"
              
            Case "�"
                Dst(I) = "D"
            Case "�"
                Dst(I) = "d"
              
            Case "�"
                Dst(I) = "E"
            Case "�"
                Dst(I) = "e"
              
            Case "�"
                Select Case PrevLetter
                    Case "�", "�", "�", "�", " "
                        Dst(I) = "Yo"
                    Case Else
                        Dst(I) = "E"
                End Select
            Case "�"
                Select Case PrevLetter
                    Case "�", "�", "�", "�", " "
                        Dst(I) = "yo"
                    Case Else
                        Dst(I) = "e"
                End Select
              
            Case "�"
                Dst(I) = "Zh"
            Case "�"
                Dst(I) = "zh"
              
            Case "�"
                Dst(I) = "Z"
            Case "�"
                Dst(I) = "z"
              
            Case "�"
                Dst(I) = "I"
            Case "�"
                Dst(I) = "i"
              
            Case "�"
                Select Case PrevLetter
                    Case "�", "�", "�", "�"
                        Dst(I - 1) = "Y"
                        Dst(I) = ""
                    Case Else
                        Dst(I) = "Y"
                End Select
            Case "�"
                Select Case PrevLetter
                    Case "�", "�", "�", "�"
                        Dst(I - 1) = "y"
                        Dst(I) = ""
                    Case Else
                        Dst(I) = "y"
                End Select
              
            Case "�"
                Dst(I) = "K"
            Case "�"
                Dst(I) = "k"
              
            Case "�"
                Dst(I) = "L"
            Case "�"
                Dst(I) = "l"
              
            Case "�"
                Dst(I) = "M"
            Case "�"
                Dst(I) = "m"
              
            Case "�"
                Dst(I) = "N"
            Case "�"
                Dst(I) = "n"
              
            Case "�"
                Dst(I) = "O"
            Case "�"
                Dst(I) = "o"
              
            Case "�"
                Dst(I) = "P"
            Case "�"
                Dst(I) = "p"
              
            Case "�"
                Dst(I) = "R"
            Case "�"
                Dst(I) = "r"
              
            Case "�"
                Dst(I) = "S"
            Case "�"
                Dst(I) = "s"
              
            Case "�"
                Dst(I) = "T"
            Case "�"
                Dst(I) = "t"
              
            Case "�"
                Dst(I) = "U"
            Case "�"
                Dst(I) = "u"
              
            Case "�"
                Dst(I) = "F"
            Case "�"
                Dst(I) = "f"
              
            Case "�"
                Dst(I) = "X"  'Kh
            Case "�"
                Dst(I) = "x"  'kh
              
            Case "�"
                Select Case PrevLetter
                    Case "�", "�"
                        Dst(I) = "S"
                    Case Else
                        Dst(I) = "Ts"
                End Select
            Case "�"
                Select Case PrevLetter
                    Case "�", "�"
                        Dst(I) = "s"
                    Case Else
                        Dst(I) = "ts"
              End Select
              
            Case "�"
                Dst(I) = "Ch"
            Case "�"
                Dst(I) = "ch"
              
            Case "�"
                Dst(I) = "Sh"
            Case "�"
                Dst(I) = "sh"
              
            Case "�"
                Dst(I) = "Sch"
            Case "�"
                Dst(I) = "sch"
              
            Case "�"
                Select Case NextLetter
                    Case "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�"
                        Dst(I) = "Y"
                    Case Else
                        Dst(I) = ""
                End Select
            Case "�"
                Select Case NextLetter
                    Case "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�"
                        Dst(I) = "y"
                    Case Else
                        Dst(I) = ""
                End Select
              
            Case "�"
                Dst(I) = "Y"
            Case "�"
                Dst(I) = "y"
              
            Case "�"
                Select Case NextLetter
                    Case "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�"
                        Dst(I) = "Y"
                    Case Else
                        Dst(I) = ""
                End Select
            Case "�"
                Select Case NextLetter
                    Case "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�"
                        Dst(I) = "y"
                    Case Else
                        Dst(I) = ""
                End Select
              
            Case "�"
                Dst(I) = "E"
            Case "�"
                Dst(I) = "e"
              
            Case "�"
                Dst(I) = "Yu"
            Case "�"
                Dst(I) = "yu"
              
            Case "�"
                Dst(I) = "Ya"
            Case "�"
                Dst(I) = "ya"
              
            '���������� �����
            Case "�"
                Dst(I) = "G"
            Case "�"
                Dst(I) = "g"
              
            Case "�"
                Dst(I) = "E"
            Case "�"
                Dst(I) = "e"
              
            Case "�"
                Select Case PrevLetter
                    Case " "
                        Dst(I) = "Yi"
                    Case Else
                        Dst(I) = "I"
                End Select
            Case "�"
                Select Case PrevLetter
                    Case " "
                        Dst(I) = "yi"
                    Case Else
                        Dst(I) = "i"
                End Select
              
            Case "�"
                Dst(I) = "I"
            Case "�"
                Dst(I) = "i"
            
            Case Else
                Dst(I) = Letter
        End Select
    Next
    
    Translit = Join(Dst, "")
End Function
