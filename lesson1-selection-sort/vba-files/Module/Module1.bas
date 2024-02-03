Attribute VB_Name = "Module1"
Sub selection_sort()
    Dim lowerBound As Integer
    Dim upperBound As Integer
    Dim index As Integer
    Dim numbers(100) As Integer

    lowerBound = 5000
    upperBound = 9000

    Range("A1:C1000").Clear

    Range("A1").Value = "Sl.No"
    Columns("A:A").EntireColumn.AutoFit
    Range("B1").Value = "Unsorted Numbers"
    Columns("B:B").EntireColumn.AutoFit
    Range("C1").Value = "Sorted Numbers"
    Columns("C:C").EntireColumn.AutoFit

    
    for index = 1 to 100
        Range("A" & (index + 1) ).Value = index

        numbers(index) = Int((upperBound - lowerBound + 1)  * Rnd()) + lowerBound
        
        Range("B" & (index + 1) ).Value = numbers(index)
    Next index

End Sub
