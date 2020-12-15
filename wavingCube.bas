Attribute VB_Name = "wavingCube"
Dim myDoc As ProductDocument
Dim myProd As Product
Dim myProds As Products

Const pi = 3.14159265358979
Public Const SIZE As Integer = 1000
Public Const MIN_SIZE As Integer = 1000
Public Const CUBES_COUNT = 15

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Sub wavingCubes()
    ' Set CATIA objects
    Set myDoc = CATIA.ActiveDocument
    Set myProd = myDoc.Product
    Set myProds = myProd.Products
    
    ' Function which modifies the size of the first cube
    setCubeSize (SIZE)
    
    ' Variables for map function (linear interpolation)
    ' Sin of any angle get values between -1, 1
    Dim minSin As Double
    minSin = -1
    Dim maxSin As Double
    maxSin = 1
    ' Desired inteval for height of the cube
    Dim minVal As Double
    minVal = 3
    Dim maxVal As Double
    maxVal = CUBES_COUNT
    
    Dim resizeFactor As Double
    
    ' Create copies of base cube
    ' The result is a matrix of cubes of size
    ' CUBES_COUNT x CUBES_COUNT
    createCubesArray (CUBES_COUNT)


    Dim nCubes As Integer
    nCubes = myProds.Count
    Dim i As Integer
    Dim distance As Double

    ' Iterate though all angles
    ' For each angle iterate all cubes
    ' There will be set an offset to the angle for each cube
    ' That offset depends on distance of cube to origin
    For Angle = 0 To 50 Step 0.1
        For i = 1 To nCubes
            
            distance = distanceToOrigin(i) 'Distance item i to origin
            angleOffset = map(distance, 0, CUBES_COUNT * SIZE, 0, 6)
            resizeFactor = map(Sin(Angle + angleOffset), minSin, maxSin, minVal, maxVal)
            
            deformCube resizeFactor, i
            
        Next
        
        ' Update working view and set a 50 milisecs delay
        myProd.Update
        DoEvents
        Sleep (50)
        
    Next

End Sub

Function distanceToOrigin(item As Integer) As Double

    Set thisCube = myProds.item(item)
    Dim pos(11)
    thisCube.Position.GetComponents pos
    x = pos(9)
    y = pos(10)
    
    distanceToOrigin = Sqr(x ^ 2 + y ^ 2)
    
End Function


Private Sub createCubesArray(items As Integer)
    Set thisCube = myProds.item(1)
    Dim refCube As Product
    Set refCube = thisCube.ReferenceProduct
    Dim initialPos(11)
    thisCube.Position.GetComponents initialPos
    Dim halfItems As Integer
    halfItems = items / 2
        For x = -halfItems To halfItems
            For y = -halfItems To halfItems
                If x = 0 And y = 0 Then
                    ' do nothing
                    ' this is to avoid create another cube at origin
                Else
                    Dim copiedCube
                    Set copiedCube = myProds.AddComponent(refCube)
                    Dim finalPos(11)
                    copiedCube.Position.GetComponents finalPos
                    finalPos(9) = initialPos(9) + x * SIZE
                    finalPos(10) = initialPos(10) + y * SIZE
                    copiedCube.Position.SetComponents finalPos
                End If
            Next
        Next
End Sub

Private Sub deformCube(mySize As Double, item As Integer)
    Set thisCube = myProds.item(item)
    Dim pos(11)
    thisCube.Position.GetComponents pos
    pos(8) = mySize
    thisCube.Position.SetComponents pos
End Sub

Private Function setCubeSize(mySize As Integer)
    Dim length1 As Length
    Set length1 = CATIA.Documents.item("Cube.CATPart").PART.Parameters.item("cubeSize")
    length1.Value = mySize
    myProd.Update
End Function
'Linear interpolation
Function map(x As Double, x1 As Double, x2 As Double, y1 As Double, y2 As Double) As Double
    y = y1 + ((x - x1) / (x2 - x1)) * (y2 - y1)
    map = y
End Function
