Module Module1
    'Assume there are 7 nodes
    Structure TreeNode

        Dim Data As String
        Dim LeftPointer As Integer
        Dim RightPointer As Integer

    End Structure

    Const NullPointer = 0
    'Note:
    ' NullPointer can start  with -1 if first Index is 0

    Dim ThisNodePtr As Integer
    Dim rootpointer As Integer
    Dim freePtr As Integer
    Dim Tree(7) As TreeNode

    Sub initialiseTree()
        Dim Index As Integer
        rootpointer = NullPointer
        freePtr = 1
        For Index = 1 To 6
            Tree(Index).LeftPointer = Index + 1
        Next
        Tree(7).LeftPointer = NullPointer
    End Sub

    Function findNode(searchItem) As Integer
        ThisNodePtr = rootpointer
        While ThisNodePtr <> NullPointer And Tree(ThisNodePtr).Data <> searchItem
            If Tree(ThisNodePtr).Data > searchItem Then
                ThisNodePtr = Tree(ThisNodePtr).LeftPointer
            Else
                ThisNodePtr = Tree(ThisNodePtr).RightPointer
            End If
        End While
        findNode = ThisNodePtr
    End Function

    Sub InsertNode(NewItem As String)
        Dim TurnedLeft As Boolean
        Dim NewNodePtr As Integer
        Dim PreviousNodePtr As Integer

        'Check if the Tree is full
        If freePtr = NullPointer Then
            Console.WriteLine("Error! Tree is full , can not insert")
        Else
            'prepare to insert
            NewNodePtr = freePtr
            freePtr = Tree(freePtr).LeftPointer
            Tree(NewNodePtr).Data = NewItem

            'set the pointers
            Tree(NewNodePtr).LeftPointer = NullPointer
            Tree(NewNodePtr).RightPointer = NullPointer

            'check if tree is empty
            If rootpointer = NullPointer Then '' insert new node at root
                rootpointer = NewNodePtr

                'find insertion point

            Else
                ThisNodePtr = rootpointer ''Start at root of the tree

                While ThisNodePtr <> NullPointer ''while not a leaf node

                    PreviousNodePtr = ThisNodePtr ''remember this node

                    If Tree(ThisNodePtr).Data > NewItem Then
                        TurnedLeft = True ''follow left pointer
                        ThisNodePtr = Tree(ThisNodePtr).LeftPointer

                    Else ''follow right pointer
                        TurnedLeft = False
                        ThisNodePtr = Tree(ThisNodePtr).RightPointer
                    End If

                End While

                If TurnedLeft = True Then
                    Tree(PreviousNodePtr).LeftPointer = NewNodePtr
                Else
                    Tree(PreviousNodePtr).RightPointer = NewNodePtr
                End If

            End If
        End If
    End Sub

    Sub outputNodes()
        Console.WriteLine("Index" & vbTab & "LeftPointer" & vbTab & "Data" & vbTab &
       "RightPointer")
        For count = 1 To 7
            Console.WriteLine(count & vbTab & Tree(count).LeftPointer & vbTab & vbTab &
           Tree(count).Data & vbTab & Tree(count).RightPointer)
        Next
    End Sub

    Sub Inorder(Thisnodeptr As Integer)
        'produces infix notation and data in ascending order
        'read left-root-right node
        If Tree(Thisnodeptr).LeftPointer <> 0 Then
            Inorder(Tree(Thisnodeptr).LeftPointer)
        End If
        Console.WriteLine(Tree(Thisnodeptr).Data)
        If Tree(Thisnodeptr).RightPointer <> 0 Then
            Inorder(Tree(Thisnodeptr).RightPointer)
        End If

    End Sub

    'Method2
    Sub Inorder2(Thisnodeptr As Integer)
        'produces infix notation and data in ascending order
        'read left-root-right node
        If Thisnodeptr <> 0 Then
            Inorder2(Tree(Thisnodeptr).LeftPointer)
            Console.WriteLine(Tree(Thisnodeptr).Data)
            Inorder2(Tree(Thisnodeptr).RightPointer)
        End If

    End Sub




    Sub Main()

        Call initialiseTree()
        Call outputNodes()
        Call InsertNode(45)
        Call InsertNode(80)
        Call InsertNode(12)
        Call InsertNode(3)
        Call InsertNode(200)
        Call InsertNode(20)
        Call InsertNode(50)
        Call Inorder2(rootpointer)
        Console.ReadKey()
        Console.ReadKey()
        Console.ReadKey()
    End Sub
End Module
