'Initialize the pagewidth and pageheight variable as double to input correctly to CreateEllipse2
Public pageWidth As Double, pageHeight As Double
'cutOff variable is the pageheight at which the holes will change from one on each side to one in each corner
'public so can be referenced from frmHoles
Public cutOff As Double
'xDistance and yDistance are distance from edge to hole center, radius & diameter refer to hole radius/diameter
'public so can be reference from frmHoles
Public xDistance As Double, yDistance As Double, diameter As Double
    
Sub screwHoles()
start:
    'Initialize the document and page
    Dim doc As document, p As Page
    'Initialize radius as double for correct input to CreateEllipse2
    Dim radius As Double

    'setting p and doc to the active page & document
    Set doc = ActiveDocument
    Set p = ActivePage

    'set document units to Millimeters
    doc.Unit = cdrMillimeter
    
    'assignment of pageWidth & pageHeight from active document
    ActivePage.GetSize pageWidth, pageHeight

    'show form to input cutOff and diameter
    frmHoles.Show

    'placeholder assignment until form is completed
    'cutOff = 50
    'diameter = 3
    'xDistance = 5
    'yDistance = 5

    'assignment of radius via diameter parameter
    radius = diameter / 2
    
    'Initialize the holes as seperate shapes (shape bottom left, shape top left etc)
    Dim sbl As Shape, stl As Shape, str As Shape, sbr As Shape, sl As Shape, sr As Shape

    
    'code segment if pageHeight is greater than cutOff value
    'draws one hole at xDistance from x edge and yDistance from y edge
    If pageHeight >= cutOff Then
        Set sbl = doc.ActiveLayer.CreateEllipse2(xDistance, _
                                                        yDistance, _
                                                        radius)
                                                        
        Set stl = doc.ActiveLayer.CreateEllipse2(xDistance, _
                                                        pageHeight - yDistance, _
                                                        radius)
                                                        
        Set str = doc.ActiveLayer.CreateEllipse2(pageWidth - xDistance, _
                                                        pageHeight - yDistance, _
                                                        radius)
                                                        
        Set sbr = doc.ActiveLayer.CreateEllipse2(pageWidth - xDistance, _
                                                        yDistance, _
                                                        radius)
    End If
            
    'code segment if pageHeight is less than cutOff value
    'draws one hole either side of label at middle of pageHeight
    If pageHeight < cutOff Then
            
        Set sl = doc.ActiveLayer.CreateEllipse2(xDistance, _
                                                        pageHeight / 2, _
                                                        radius)
        Set sr = doc.ActiveLayer.CreateEllipse2(pageWidth - xDistance, _
                                                        pageHeight / 2, _
                                                        radius)
            
    End If
    
End Sub
