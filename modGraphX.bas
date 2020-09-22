Attribute VB_Name = "modGraphX"
'--------------------------------------------------------
'       Copyright 2002, Anoop Sankar
'You may freely use, modify and distribute this source
'code, provided that you do not remove this message.
'But, you are NOT allowed to distribute the compiled
'version (.EXE,.DLL,.OCX etc etc.) of this program
'or any program which uses the below code without my
'consent.
'
'If you modified something, put your name below..
'
'Orginal Code : Anoop Sankar (anoops@gmx.net)
'Modified by  : No one so far
'
'Last Update : Oct 3,2002
'Visit www.smilehouse.cjb.net for more source code
'-------------------------------------------------------
'
'The ShiftColor function was taken from gonchuki's
'Chameleon buttons source.
'The rest of the stuff was done by me.
'
'And oh.. the photograph was downloaded from the net.
'The beautiful face is that of Aishwarya Rai, a popular
'Indian actress, who was also Miss World in 1993.
'
'-------------------------------------------------------

Public Sub Interleave(Source As Object, Destination As Object, Gap As Long, Color As Long, Optional Effect As Integer = 1)
        
    'Interleave Routine
    
    'Effects    :   1 = Vertical
    '               2 = Horizontal
    '               3 = Both
    '               4 = Color Tint \ Blend
    
    Dim i As Integer, j As Integer, srcColor As Long
    
    'set background color to the color of the interleave
    Destination.BackColor = Color
    
    'the two For loops are used to browse through every
    'single pixel of the image
    For i = 0 To Source.Width
        For j = 0 To Source.Height
            'get the color of the current pixel
            srcColor = Source.Point(i, j)
                        
            'The basic logic in the below effects is that
            'you draw a pixel only when it comes at a
            'co-ordinate that is a multiple of 'gap'
                        
            Select Case Effect
                
                Case 1
                    If i Mod Gap = 0 Then Destination.PSet (i, j), srcColor
                Case 2
                    If j Mod Gap = 0 Then Destination.PSet (i, j), srcColor
                Case 3
                    If j Mod Gap = 0 Or i Mod Gap = 0 Then Destination.PSet (i, j), srcColor
                Case 4
                    If i Mod Gap = 0 Xor j Mod Gap = 0 Then Destination.PSet (i, j), srcColor
                
            End Select
            
        Next j
        'If do not want to see the program at work, comment
        'the below line
        DoEvents
    
    Next i
   
End Sub

Public Sub Tarnish(Source As Object, Destination As Object, Level As Integer)
    Randomize Timer
    
    'Tarnish the image
    
    Dim i As Integer, j As Integer, r As Integer, srcColor As Long
    
    'clear the background
    Destination.Cls
    
    'browse through pixel
    For i = 0 To Source.Width
        For j = 0 To Source.Height
            'get pixel color
            srcColor = Source.Point(i, j)
            'get a rnd # between 1 and 'Level'
            r = Int(Rnd * Level) + 1
            'change srcColor randomly. More the level,
            'the more the orginal and applied pixels differ.
            srcColor = Int(srcColor / r)
            'draw the pixel onto the destination
            Destination.PSet (i, j), srcColor
        Next j
        DoEvents
    Next i

End Sub

Public Sub Churn(Source As Object, Destination As Object, Level As Integer)
    
    'silly name, ha?
    'suggest a good one if you can! ;)
    
    'uses a similar technique as tarnish,
    'but the effect is different
    
    Randomize Timer
    
    Dim i As Integer, j As Integer, r As Integer, srcColor As Long
    
    'clear background
    Destination.Cls

    For i = 0 To Source.Width
        For j = 0 To Source.Height
            srcColor = Source.Point(i, j)
            
            r = Int(Rnd * Level) + 1
            
            srcColor = ShiftColor(srcColor, r)
                        
            Destination.PSet (i, j), srcColor
        Next j
        DoEvents
    Next i

End Sub

Public Sub PencilDraw(Source As Object, Destination As Object, Level As Long)
    
    'not really a suitable name
    '
    'The level variable controls almost
    'everything here. Values of 5, 100,500
    'all give totally different results!
    '
    
    Dim i As Integer, j As Integer, srcColor As Long
    Dim Base As Long
            
    Destination.Cls

    'The technique here is to round off all pixel colors to
    'the nearest multiple of 'Base', lower than its orginal color
    '(Pretty similar to the keyword Int.)
    
    Base = &HFFFFFF \ Level
    
    For i = 0 To Source.Width
        For j = 0 To Source.Height
            srcColor = Source.Point(i, j)
            
            'rounding off
            srcColor = Base * (srcColor \ Base)
            'if you are wondering whether Base and Base
            'cancel out in the previous expression, notice
            'the slash used for division. Backward slash
            'gives you only the integer part and no decimals
                               
            Destination.PSet (i, j), srcColor
        Next j
        DoEvents
    Next i
End Sub

Public Sub Negative(Source As Object, Destination As Object)
    Dim i As Integer, j As Integer, srcColor As Long
        
    'negative of a photo? No probs
        
    For i = 0 To Source.Height
        For j = 0 To Source.Width
            srcColor = Source.Point(j, i)
            
            'negative
            'didn't realize that was this simple until now!
            
            srcColor = &HFFFFFF - srcColor
                                    
            Destination.PSet (j, i), srcColor
      
        Next j
        DoEvents
     Next i
    
End Sub

Public Sub GrayScale(Source As Object, Destination As Object, Optional Channel As Integer = 1, Optional OldPhoto As Boolean = False)
    Dim i As Integer, j As Integer, srcColor As Long
    Dim Red As Long, Blue As Long, Green As Long
    Dim chColor As Long
        
    'GrayScale
            
    'Channel:   1 = Green
    '           2 = Blue
    '           3 = Red
            
    For i = 0 To Source.Height
        For j = 0 To Source.Width
            srcColor = Source.Point(j, i)
            
            'get each component of color
            Green = ((srcColor \ &H100) Mod &H100)
            Red = (srcColor And &HFF)
            Blue = ((srcColor \ &H10000) Mod &H100)
           
            Select Case Channel
                
                Case 1
                    chColor = Green
                Case 2
                    chColor = Blue
                Case 2
                    chColor = Red
            End Select
           
            'What I did below was, replace two color components
            'component with the third one,'Channel'.
            'The real grayscaling must be done equally
            'to all colors and I don't know any easy way
            'to do this.
            'But atleast this works! :-)
            
            srcColor = RGB(chColor, chColor, chColor)
            
            If OldPhoto Then
                srcColor = RGB(Green, Green, Blue)
            End If
            
            'the below lines uses my orginal approach.
            'it does create a grayscale image, but not a
            'very good one.. not 'bright' enough. stupid
            'algo, but just included it here anyway!
            
            'r1 = Hex(r)
            'Debug.Print Hex(srcColor)
            'clstr = "&H8" & r1
            
            'srcColor = RGB(Val(clstr), Val(clstr), Val(clstr))
            
            'If srcColor > &H808080 Then
            '    srcColor = (srcColor / 1.9998) + &H808080
            'Else
            '    srcColor = &H808080 - (srcColor / 1.9998)
            'End If
                                    
            Destination.PSet (j, i), srcColor
      
        Next j
        DoEvents
     Next i
    
End Sub

Public Sub TriColor(Source As Object, Destination As Object)
    'draws a tricolor band blend ..
    'uses the colors of the flag of India => Saffron, White and Green

    Destination.Cls

    Dim i As Integer, j As Integer, k As Integer, srcColor As Long
    Dim chColor As Long

    'three bands
    For k = 1 To 3
        
        'each band has different color
        Select Case k
            Case 1: chColor = RGB(&HEC, &H53, 0) 'saffron(ish)
            Case 2: chColor = vbWhite
            Case 3: chColor = vbGreen
        End Select
        
        'draw each band
        For i = (k - 1) * (Source.Height / 3) To k * (Source.Height / 3)
            For j = 0 To Source.Width
                srcColor = Source.Point(j, i)
                Destination.PSet (j, i), srcColor
                
                'use the color blend code
                If i Mod 2 = 0 Xor j Mod 2 = 0 Then Destination.PSet (j, i), chColor
            Next j
        DoEvents
        Next i

    Next k
End Sub


Public Sub Mosaic(Source As Object, Destination As Object, Gap As Integer)
    
    ' not really neat
    ' bad logic, something is missing somewhere
        
    Dim PntArr(500, 500)
    
    For i = 0 To Source.Width
        For j = 0 To Source.Height Step Gap
            srcColor = Source.Point(i, j)
            For k = j To j + (Gap - 1)
                PntArr(i, k) = srcColor
            Next k
        Next j
     Next i
    
    
    For i = 0 To Source.Width
        For j = 0 To Source.Height
            Destination.PSet (i, j), PntArr(i, j)
        Next j
        DoEvents
    Next i
    
End Sub

Private Function ShiftColor(ByVal Color As Long, ByVal Value As Long, Optional isXP As Boolean = False) As Long
'shift color code by gonchuki
'taken from Chameleon buttons (PSC 2001)

'this function will add or remove a certain color
'quantity and return the result

Dim Red As Long, Blue As Long, Green As Long

If isXP = False Then
    Blue = ((Color \ &H10000) Mod &H100) + Value
Else
    Blue = ((Color \ &H10000) Mod &H100)
    Blue = Blue + ((Blue * Value) \ &HC0)
End If
Green = ((Color \ &H100) Mod &H100) + Value
Red = (Color And &HFF) + Value
    
    'check red
    If Red < 0 Then
        Red = 0
    ElseIf Red > 255 Then
        Red = 255
    End If
    'check green
    If Green < 0 Then
        Green = 0
    ElseIf Green > 255 Then
        Green = 255
    End If
    'check blue
    If Blue < 0 Then
        Blue = 0
    ElseIf Blue > 255 Then
        Blue = 255
    End If

ShiftColor = RGB(Red, Green, Blue)
End Function

