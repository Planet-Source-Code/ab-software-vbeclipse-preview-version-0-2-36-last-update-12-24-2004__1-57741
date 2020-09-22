VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HashTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'----------------------------------------------
' HASHTABLE class module
'
' This class implements a hashtable, a structure that offers many
' of the features of a collectior or dictionary, and is often
' even faster than the built-in collection.
'
' NOTE: must make Item the default member, using the Tools | Procedure
' Attributes dialog
'
' Usage:
'   Dim ht As New HashTable
'   ht.SetSize 10000           ' initial number of slots (the higher,
'  the better)
'
'   ' enforce case-insensitive key search
'   ht.IgnoreCase = True
'   ' add values
'   ht.Add "key", value        ' add a value associated to a key
'   ' count how many values are in the table
'   Print ht.Count
'   ' read/write a value
'   Print ht("key")
'   ht("key") = newValue
'
'   ' remove a value
'   ht.Remove "key"
'   ' remove all values
'   ht.RemoveAll
'   ' check whether a value exists
'   If ht.Exists("key") Then ...
'
'   ' get the array of keys and values
'   Dim keys() As String, values() As String
'   keys() = ht.Keys
'   values() = ht.Values
'
'----------------------------------------------

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As _
    Any, Source As Any, ByVal Bytes As Long)

' default values
Const DEFAULT_HASHSIZE As Long = 1024
Const DEFAULT_LISTSIZE As Long = 2048
Const DEFAULT_CHUNKSIZE As Long = 1024

Private Type SlotType
    Key As String
    Value As Object
    NextItem As Long      ' 0 if last item
End Type

' for each hash code this array holds the first element
' in slotTable() with the corresponding hash code
Dim hashTbl() As Long
' the array that holds the data
Dim slotTable() As SlotType

' pointer to first free slot
Dim FreeNdx As Long

' size of hash table
Dim m_HashSize As Long
' size of slot table
Dim m_ListSize As Long
' chunk size
Dim m_ChunkSize As Long
' items in the slot table
Dim m_Count As Long

' member variable for IgnoreCase property
Private m_IgnoreCase As Boolean
Private m_ReplaceableKeys As Boolean
' True if keys are searched in case-unsensitive mode
' this can be assigned to only when the hash table is empty

Property Get IgnoreCase() As Boolean
    IgnoreCase = m_IgnoreCase
End Property

Property Let IgnoreCase(ByVal newValue As Boolean)
    If m_Count Then
        Err.Raise 1001, , "The Hash Table isn't empty"
    End If
    m_IgnoreCase = newValue
End Property

Property Get ReplaceableKeys() As Boolean
    ReplaceableKeys = m_ReplaceableKeys
End Property

Property Let ReplaceableKeys(ByVal New_ReplaceableKeys As Boolean)
    m_ReplaceableKeys = New_ReplaceableKeys
End Property

' initialize the hash table

Sub SetSize(ByVal HashSize As Long, Optional ByVal ListSize As Long, _
    Optional ByVal ChunkSize As Long)
    ' provide defaults
    If ListSize <= 0 Then ListSize = m_ListSize
    If ChunkSize <= 0 Then ChunkSize = m_ChunkSize
    ' save size values
    m_HashSize = HashSize
    m_ListSize = ListSize
    m_ChunkSize = ChunkSize
    m_Count = 0
    ' rebuild tables
    FreeNdx = 0
    ReDim hashTbl(0 To HashSize - 1) As Long
    ReDim slotTable(0) As SlotType
    ExpandSlotTable m_ListSize
End Sub

' check whether an item is in the hash table

Function Exists(Key As String) As Boolean
    Exists = GetSlotIndex(Key) <> 0
End Function

' add a new element to the hash table

Sub Add(Key As String, Value As Object)
    Dim ndx As Long, Create As Boolean
    
    ' get the index to the slot where the value is
    ' (allocate a new slot if necessary)
    Create = True
    ndx = GetSlotIndex(Key, Create)
    
    If Create Or m_ReplaceableKeys Then
        ' the item was actually added
        If IsObject(Value) Then
            Set slotTable(ndx).Value = Value
        Else
            slotTable(ndx).Value = Value
        End If
    Else
        ' raise error "This key is already associated with an item of this
        ' collection"
        Err.Raise 457
    End If
End Sub

' the value associated to a key
' (empty if not found)

Property Get Item(ByVal Key As String, Optional ByVal default As Object) As Object
    Dim ndx As Long
    ' get the index to the slot where the value is
    ndx = GetSlotIndex(Key)
    If ndx = 0 Then
        ' return Empty if not found
    ElseIf IsObject(slotTable(ndx).Value) Then
        Set Item = slotTable(ndx).Value
    Else
        Item = slotTable(ndx).Value
    End If
    
    If Not IsObject(Item) Then
       If Len(Item) = 0 Then
          If IsMissing(default) Then
             Item = vbNullString
          'ElseIf default Is Nothing Then
          '   Item = vbNullString
          Else
             Item = default
          End If
       End If
    Else
       If Item Is Nothing Then Set Item = default
    End If
    
End Property

'Property Let Item(Key As String, value As String)
'    Dim ndx As Long
    ' get the index to the slot where the value is
    ' (allocate a new slot if necessary)
'    ndx = GetSlotIndex(Key, True)
    ' store the value
'    slotTable(ndx).value = value
'End Property

' remove an item from the hash table

Sub Remove(Key As String)
    Dim ndx As Long, HCode As Long, LastNdx As Long
    ndx = GetSlotIndex(Key, False, HCode, LastNdx)
    ' raise error if no such element
    If ndx = 0 Then Err.Raise 5
    
    If LastNdx Then
        ' this isn't the first item in the slotTable() array
        slotTable(LastNdx).NextItem = slotTable(ndx).NextItem
    ElseIf slotTable(ndx).NextItem Then
        ' this is the first item in the slotTable() array
        ' and is followed by one or more items
        hashTbl(HCode) = slotTable(ndx).NextItem
    Else
        ' this is the only item in the slotTable() array
        ' for this hash code
        hashTbl(HCode) = 0
    End If
    
    ' put the element back in the free list
    slotTable(ndx).NextItem = FreeNdx
    FreeNdx = ndx
    ' we have deleted an item
    m_Count = m_Count - 1
    
End Sub

' remove all items from the hash table

Sub RemoveAll()
    SetSize m_HashSize, m_ListSize, m_ChunkSize
End Sub

' the number of items in the hash table

Property Get Count() As Long
    Count = m_Count
End Property

' the array of all keys
' (VB5 users: convert return type to string)

Property Get Keys(Optional ByVal Sorted As Boolean = True, _
                  Optional ByVal Desc As Boolean = False) As Variant
    Dim i As Long, ndx As Long
    Dim n As Long
    
    If m_Count = 0 Then Exit Property
    
    ReDim res(0 To m_Count - 1) As String
    
    For i = 0 To m_HashSize - 1
        ' take the pointer from the hash table
        ndx = hashTbl(i)
        ' walk the slottable() array
        Do While ndx
            res(n) = slotTable(ndx).Key
            n = n + 1
            ndx = slotTable(ndx).NextItem
        Loop
    Next
        
    ' sort values before returning
    If Sorted = True Then
       Sort res, , Desc
    End If
    
    ' assign to the result
    Keys = res()
End Property

' the array of all values
' (VB5 users: convert return type to String)

Property Get Values(Optional ByVal Sorted As Boolean = True, _
                    Optional ByVal Desc As Boolean = False) As Variant
    Dim i As Long, ndx As Long
    Dim n As Long
    
    If m_Count = 0 Then Exit Property

    ReDim res(0 To m_Count - 1) As Object
    
    For i = 0 To m_HashSize - 1
        ' take the pointer from the hash table
        ndx = hashTbl(i)
        ' walk the slottable() array
        Do While ndx
            If IsObject(slotTable(ndx).Value) Then
               Set res(n) = slotTable(ndx).Value
            Else
               res(n) = slotTable(ndx).Value
            End If
            n = n + 1
            ndx = slotTable(ndx).NextItem
        Loop
    Next
        
    ' sort values before return
    'If Sorted = True Then
    '   Sort res(), , Desc
    'End If
    
    ' assign to the resulting
    Values = res()
End Property

'-----------------------------------------
' Private procedures
'-----------------------------------------

' QuickSort an array of any type
' QuickSort is especially convenient with large arrays (>1,000
' items) that contains items in random order. Its performance
' quickly degrades if the array is already almost sorted. (There are
' variations of the QuickSort algorithm that work good with
' nearly-sorted arrays, though, but this routine doesn't use them.)
'
' NUMELS is the index of the last item to be sorted, and is
' useful if the array is only partially filled.
'
' Works with any kind of array, except UDTs and fixed-length
' strings, and including objects if your are sorting on their
' default property. String are sorted in case-sensitive mode.
'
' You can write faster procedures if you modify the first two lines
' to account for a specific data type, eg.
' Sub QuickSortS(arr() As Single, Optional numEls As Variant, Optional descending As Boolean)
'   Dim value As Single, temp As Single
Private Sub Sort(arr() As String, Optional numEls As Variant, Optional Descending As Boolean)

    Dim Value As String, Temp As String
    Dim sp As Integer
    Dim leftStk(32) As Long, rightStk(32) As Long
    Dim leftNdx As Long, rightNdx As Long
    Dim i As Long, j As Long

    ' account for optional arguments
    If IsMissing(numEls) Then numEls = UBound(arr)
    ' init pointers
    leftNdx = LBound(arr)
    rightNdx = numEls
    ' init stack
    sp = 1
    leftStk(sp) = leftNdx
    rightStk(sp) = rightNdx

    Do
        If rightNdx > leftNdx Then
            Value = arr(rightNdx)
            i = leftNdx - 1
            j = rightNdx
            ' find the pivot item
            If Descending Then
                Do
                    Do: i = i + 1: Loop Until arr(i) <= Value
                    Do: j = j - 1: Loop Until j = leftNdx Or arr(j) >= Value
                    Temp = arr(i)
                    arr(i) = arr(j)
                    arr(j) = Temp
                Loop Until j <= i
            Else
                Do
                    Do: i = i + 1: Loop Until arr(i) >= Value
                    Do: j = j - 1: Loop Until j = leftNdx Or arr(j) <= Value
                    Temp = arr(i)
                    arr(i) = arr(j)
                    arr(j) = Temp
                Loop Until j <= i
            End If
            ' swap found items
            Temp = arr(j)
            arr(j) = arr(i)
            arr(i) = arr(rightNdx)
            arr(rightNdx) = Temp
            ' push on the stack the pair of pointers that differ most
            sp = sp + 1
            If (i - leftNdx) > (rightNdx - i) Then
                leftStk(sp) = leftNdx
                rightStk(sp) = i - 1
                leftNdx = i + 1
            Else
                leftStk(sp) = i + 1
                rightStk(sp) = rightNdx
                rightNdx = i - 1
            End If
        Else
            ' pop a new pair of pointers off the stacks
            leftNdx = leftStk(sp)
            rightNdx = rightStk(sp)
            sp = sp - 1
            If sp = 0 Then Exit Do
        End If
    Loop
End Sub

Private Sub Class_Initialize()
    ' initialize the tables at default size
    SetSize DEFAULT_HASHSIZE, DEFAULT_LISTSIZE, DEFAULT_CHUNKSIZE
    
    m_ReplaceableKeys = True
    m_IgnoreCase = False
End Sub

' expand the slotTable() array

Private Sub ExpandSlotTable(ByVal numEls As Long)
    Dim newFreeNdx As Long, i As Long
    newFreeNdx = UBound(slotTable) + 1
    
    ReDim Preserve slotTable(0 To UBound(slotTable) + numEls) As SlotType
    ' create the linked list of free items
    For i = newFreeNdx To UBound(slotTable)
        slotTable(i).NextItem = i + 1
    Next
    ' overwrite the last (wrong) value
    slotTable(UBound(slotTable)).NextItem = FreeNdx
    ' we now know where to pick the first free item
    FreeNdx = newFreeNdx
End Sub

' return the hash code of a string

Private Function HashCode(Key As String) As Long
    Dim lastEl As Long, i As Long
    
    ' copy ansi codes into an array of long
    lastEl = (Len(Key) - 1) \ 4
    ReDim codes(lastEl) As Long
    ' this also converts from Unicode to ANSI
    CopyMemory codes(0), ByVal Key, Len(Key)
    
    ' XOR the ANSI codes of all characters
    For i = 0 To lastEl
        HashCode = HashCode Xor codes(i)
    Next
    
End Function

' get the index where an item is stored or 0 if not found
' if Create = True the item is created
'
' on exit Create=True only if a slot has been actually created

Private Function GetSlotIndex(ByVal Key As String, Optional Create As Boolean, _
    Optional HCode As Long, Optional LastNdx As Long) As Long
    Dim ndx As Long
    
    ' raise error if invalid key
    If Len(Key) = 0 Then Exit Function 'Err.Raise 1001, , "Invalid key"
    
    ' keep case-unsensitiveness into account
    If m_IgnoreCase Then Key = UCase$(Key)
    ' get the index in the hashTbl() array
    HCode = HashCode(Key) Mod m_HashSize
    ' get the pointer to the slotTable() array
    ndx = hashTbl(HCode)
    
    ' exit if there is no item with that hash code
    Do While ndx
        ' compare key with actual value
        If slotTable(ndx).Key = Key Then Exit Do
        ' remember last pointer
        LastNdx = ndx
        ' check the next item
        ndx = slotTable(ndx).NextItem
    Loop
    
    ' create a new item if not there
    If ndx = 0 And Create Then
        ndx = GetFreeSlot()
        PrepareSlot ndx, Key, HCode, LastNdx
    Else
        ' signal that no item has been created
        Create = False
    End If
    ' this is the return value
    GetSlotIndex = ndx

End Function

' return the first free slot

Private Function GetFreeSlot() As Long
    ' allocate new memory if necessary
    If FreeNdx = 0 Then ExpandSlotTable m_ChunkSize
    ' use the first slot
    GetFreeSlot = FreeNdx
    ' update the pointer to the first slot
    FreeNdx = slotTable(GetFreeSlot).NextItem
    ' signal this as the end of the linked list
    slotTable(GetFreeSlot).NextItem = 0
    ' we have one more item
    m_Count = m_Count + 1
End Function

' assign a key and value to a given slot

Private Sub PrepareSlot(ByVal Index As Long, ByVal Key As String, _
    ByVal HCode As Long, ByVal LastNdx As Long)
    ' assign the key
    ' keep case-sensitiveness into account
    If m_IgnoreCase Then Key = UCase$(Key)
    slotTable(Index).Key = Key
    
    If LastNdx Then
        ' this is the successor of another slot
        slotTable(LastNdx).NextItem = Index
    Else
        ' this is the first slot for a given hash code
        hashTbl(HCode) = Index
    End If
End Sub

