Attribute VB_Name = "LaoWrapMacro"
' ============================================================
' Lao line-wrap helper for Microsoft Word
' phetsarath.mts.la - SIL OFL / free to use
'
' FixLaoWrap  : inserts invisible zero-width spaces (U+200B) at
'               Lao syllable boundaries in the selection (or the
'               whole document if nothing is selected) so older
'               versions of Word can wrap long Lao sentences.
'               The characters are INVISIBLE and do not print.
' RemoveLaoWrap : removes all zero-width spaces again.
'
' Modern Microsoft 365 usually wraps Lao correctly by itself -
' first try: File > Options > Language > "Office authoring
' languages and proofing" > Add a language > Lao, restart Word,
' then select the text > Review > Language > Set Proofing
' Language > Lao. Use this macro only if wrapping still fails.
' ============================================================
Option Explicit

Private Function IsLao(ByVal code As Long) As Boolean
    IsLao = (code >= &HE80 And code <= &HEFF)
End Function

Private Function IsPreVowel(ByVal code As Long) As Boolean
    ' pre-vowels always start a new syllable
    IsPreVowel = (code = &HEC0 Or code = &HEC1 Or code = &HEC2 _
               Or code = &HEC3 Or code = &HEC4)
End Function

Private Function IsConsonant(ByVal code As Long) As Boolean
    IsConsonant = (code >= &HE81 And code <= &HEAE) _
               Or code = &HEDC Or code = &HEDD
End Function

Private Function IsVowelSign(ByVal code As Long) As Boolean
    ' dependent vowels & tone marks that BELONG to the previous initial
    IsVowelSign = (code >= &HEB0 And code <= &HEBD) _
               Or (code >= &HEC6 And code <= &HECD)
End Function

Private Function IsClusterSecond(ByVal code As Long) As Boolean
    ' second element of an initial cluster (after ho or others)
    IsClusterSecond = (code = &HEBC Or code = &HEA7 Or code = &HEA5 _
                    Or code = &HEA3 Or code = &HE8D Or code = &HE99 _
                    Or code = &HEA1 Or code = &HE87)
End Function

Public Sub FixLaoWrap()
    Dim rng As Range
    If Selection.Type = wdSelectionIP Then
        Set rng = ActiveDocument.Content
    Else
        Set rng = Selection.Range
    End If
    Dim s As String: s = rng.Text
    Dim out As String: out = ""
    Dim i As Long, n As Long: n = Len(s)
    Dim prev As Long, cur As Long, nxt As Long
    For i = 1 To n
        cur = AscW(Mid$(s, i, 1))
        prev = 0: If i > 1 Then prev = AscW(Mid$(s, i - 1, 1))
        nxt = 0: If i < n Then nxt = AscW(Mid$(s, i + 1, 1))
        If IsLao(prev) And prev <> &H200B Then
            If IsPreVowel(cur) Then
                out = out & ChrW(&H200B)
            ElseIf IsConsonant(cur) And Not IsPreVowel(prev) _
                   And (IsVowelSign(nxt) Or IsPreVowel(nxt) _
                   Or (cur = &HEAB And IsClusterSecond(nxt))) Then
                ' consonant starting a new syllable - but never right
                ' after a pre-vowel (it must stay glued to its consonant)
                ' and not as the second element of an initial cluster
                If Not (IsClusterSecond(cur) And IsConsonant(prev)) Then
                    out = out & ChrW(&H200B)
                End If
            End If
        End If
        out = out & Mid$(s, i, 1)
    Next i
    If out <> s Then rng.Text = out
    MsgBox "ສຳເລັດ! ໄດ້ເພີ່ມຈຸດຕັດຄຳທີ່ເບິ່ງບໍ່ເຫັນແລ້ວ" & vbCrLf & _
           "(Invisible break points inserted.)", vbInformation
End Sub

Public Sub RemoveLaoWrap()
    Dim rng As Range
    If Selection.Type = wdSelectionIP Then
        Set rng = ActiveDocument.Content
    Else
        Set rng = Selection.Range
    End If
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ChrW(&H200B)
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceAll
    End With
    MsgBox "ລຶບຈຸດຕັດຄຳອອກແລ້ວ (zero-width spaces removed).", vbInformation
End Sub
