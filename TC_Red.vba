Sub tracked_to_bluefont()
    tempState = ActiveDocument.TrackRevisions
    ActiveDocument.TrackRevisions = False
    For Each Change In ActiveDocument.Revisions
        Set myRange = Change.Range
        myRange.Revisions.AcceptAll
        myRange.Font.Color = 00255
    Next
    ActiveDocument.TrackRevisions = tempState
End Sub
