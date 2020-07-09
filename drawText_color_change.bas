Attribute VB_Name = "drawText_color_change"
'catvba  drawText_color_change using-'ver0.1.0'  by Kantoku
'Sample that changes color with text size ver0.0.2

Option Explicit

Private Const PROPCOLOR = CatTextProperty.catColor
Private Const PROPSIZE = CatTextProperty.catFontSize
Private Const OTHERCOLOR = "0, 0, 0"
Private Const OTHERSUBCOLOR = 255

Private p_PropColorMap As Variant

Sub CATMain()
    
    If Not CanExecute("DrawingDocument") Then Exit Sub
    
    Dim txts As Collection
    Set txts = getShowTxts()
    If txts Is Nothing Then Exit Sub
    
    Dim colorMap As Variant
    colorMap = initColorMap()
    
    Dim sizeDic As Object
    Set sizeDic = groupBySize(txts)
    
    Call execChangeColor(sizeDic, colorMap)
    
    MsgBox "Done"
End Sub

Private Sub execChangeSubColor( _
    ByVal dt As DrawingText)
    
    Dim max As Long
    max = Len(dt.Text) + 1
    
    Dim color As Long, key As Long
    
    Dim st As Long, cnt As Long, num As Long
    st = 1
    cnt = 1
    
    Dim res As Long
    Do
        res = dt.GetParameterOnSubString(PROPSIZE, st, cnt)
        
        If res <> 0 And max > st + cnt Then
            cnt = cnt + 1
            GoTo continue
        End If
        
        'hit
        If res = 0 Then
            num = cnt - 1
        Else
            num = cnt
        End If
        
        key = CLng(dt.GetParameterOnSubString( _
            PROPSIZE, st, num) / 1000)
            
        If UBound(p_PropColorMap) > key Then
            'In colormap
            color = p_PropColorMap(key)
        Else
            'other
            color = OTHERSUBCOLOR
        End If
            
        Call dt.SetParameterOnSubString( _
            PROPCOLOR, st, num, color)
    
        If cnt = 1 Then
            cnt = cnt + 1
        End If
        
        st = st + cnt - 1
        cnt = 1

continue:
        If max < st + cnt Then Exit Do
        
    Loop
        
End Sub

Private Sub execChangeColor( _
    ByVal group As Object, _
    ByVal colorMap As Variant)
    
    Dim sel As selection
    Set sel = CATIA.ActiveDocument.selection
    
    Dim vis As VisPropertySet
    Set vis = sel.VisProperties
    
    CATIA.HSOSynchronized = False
    
    Dim key As Variant ' Long
    Dim rgbAry As Variant
    Dim rgbTxt As String
    Dim dt As DrawingText
    For Each key In group.keys

        If key < 0 Then
        
            If IsEmpty(p_PropColorMap) Then
                Call setPropColorMap(colorMap)
            End If
            
            'sub color
            For Each dt In group(key)
                Call execChangeSubColor(dt)
            Next
            
            GoTo continue
        End If

        If UBound(colorMap) > key Then
            'In colormap
            rgbTxt = colorMap(key)
        Else
            'other
            rgbTxt = OTHERCOLOR
        End If

        rgbAry = Split(rgbTxt, ",")
        
        sel.Clear
        For Each dt In group(key)
            sel.Add dt
        Next
        
        Call vis.SetRealColor( _
            CLng(rgbAry(0)), _
            CLng(rgbAry(1)), _
            CLng(rgbAry(2)), _
            1)
            
continue:
    Next
    
    sel.Clear
    
    CATIA.HSOSynchronized = True
    
End Sub

'return dic(txtsize,lst(drawtxt))
Private Function groupBySize( _
    ByVal txts As Collection) _
    As Object

    Dim dic As Object
    Set dic = initDic()
    
    Dim dt As DrawingText
    Dim key As Long, subsize As Long
    Dim lst As Collection
    
    For Each dt In txts
        subsize = dt.GetParameterOnSubString(PROPSIZE, 0, 0)
        
        If subsize = 0 Then
            key = -1
        Else
            key = CLng(subsize / 1000)
        End If
        
        If dic.Exists(key) Then
            Call dic(key).Add(dt)
        Else
            Set lst = New Collection
            lst.Add dt
            Call dic.Add(key, lst)
        End If
    Next
    
    Set groupBySize = dic
    
End Function

Private Function initDic() _
    As Object
    
    Set initDic = CreateObject("Scripting.Dictionary")
    
End Function

Private Function getShowTxts() _
    As Collection
    
    Set getShowTxts = Nothing
    
    Dim vr As Viewer
    Set vr = CATIA.ActiveWindow.ActiveViewer
    
    Dim vp2D As Object 'Viewer2D
    Set vp2D = vr.Viewpoint2D
    
    Dim state_zoom As Double
    state_zoom = vp2D.Zoom
    
    vp2D.Zoom = 0.0000001
    'vr.Update
    
    Dim sel As selection
    Set sel = CATIA.ActiveDocument.selection

    CATIA.HSOSynchronized = False
    
    sel.Clear
    sel.Search "CATDrwSearch.DrwText,scr"

    If sel.count < 1 Then
        MsgBox "Text not found"
        Exit Function
    End If
    
    Dim lst As Collection
    Set lst = New Collection
    
    Dim i As Long
    For i = 1 To sel.count
        lst.Add sel.Item2(i).value
    Next
    
    sel.Clear
    
    CATIA.HSOSynchronized = True
    
    vp2D.Zoom = state_zoom
    'vr.Update
    
    Set getShowTxts = lst
    
End Function

Private Sub setPropColorMap( _
    ByVal colorMap As Variant)
    
    Dim doc As DrawingDocument
    Set doc = CATIA.ActiveDocument
    
    Dim txts As DrawingTexts
    Set txts = doc.Sheets.ActiveSheet.views.ActiveView.Texts
    
    Dim txt As DrawingText
    Set txt = txts.Add("hoge", 0, 0)
    
    Dim sel As selection
    Set sel = doc.selection
    
    sel.Clear
    sel.Add txt
    
    Dim vis As VisPropertySet
    Set vis = sel.VisProperties
    
    Dim rgbAry As Variant
    Dim res As Long
    
    Dim ary() As Long
    ReDim ary(UBound(colorMap))
    Dim i As Long
    For i = 0 To UBound(colorMap)
        rgbAry = Split(colorMap(i), ",")
        Call vis.SetRealColor( _
            CLng(rgbAry(0)), _
            CLng(rgbAry(1)), _
            CLng(rgbAry(2)), _
            1)
        
        res = txt.GetParameterOnSubString(PROPCOLOR, 0, 0)
        ary(i) = res
    Next
    
    sel.Delete
    
    p_PropColorMap = ary
    
End Sub

Private Function initColorMap() _
    As Variant 'array(str)
    
    initColorMap = Array( _
        "255, 255, 0", _
        "128, 0, 255", _
        "211, 178, 125", _
        "255, 128, 0", _
        "0, 255, 255", _
        "255, 0, 0", _
        "0, 0, 255", _
        "0, 128, 255", _
        "0, 255, 0", _
        "0, 128, 0", _
        "255, 0, 255", _
        "128, 64, 64")

End Function

