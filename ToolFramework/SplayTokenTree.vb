Imports System.Collections.Generic
Imports Microsoft.Nautilus.Text
Imports SuperClassifier

Friend Class SplayTokenTree
    Private count As Integer
    Private Shared header As New BinaryNode(Nothing)

    Friend Property Root As BinaryNode

    Friend Sub Insert(token As Token)
        If Root Is Nothing Then
            count = 1
            Root = New BinaryNode(token)
            Return
        End If

        Splay(token.TokenStart)
        If token.TokenStart <> Root.Token.TokenStart Then
            Dim binaryNode1 As New BinaryNode(token)
            If token.TokenStart < Root.Token.TokenStart Then
                binaryNode1.LeftChild = Root.LeftChild
                binaryNode1.RightChild = Root
                Root.LeftChild = Nothing
            Else
                binaryNode1.RightChild = Root.RightChild
                binaryNode1.LeftChild = Root
                Root.RightChild = Nothing
            End If
            Root = binaryNode1
            count += 1
        End If
    End Sub

    Private Sub Splay(tokenStart As Integer)
        Dim binaryNode1 As BinaryNode = header
        Dim binaryNode2 As BinaryNode = header
        Dim binaryNode3 As BinaryNode = Root
        header.LeftChild = Nothing
        header.RightChild = Nothing

        While True
            If tokenStart < binaryNode3.Token.TokenStart Then
                If binaryNode3.LeftChild Is Nothing Then
                    Exit While
                End If

                If tokenStart < binaryNode3.LeftChild.Token.TokenStart Then
                    Dim leftChild1 As BinaryNode = binaryNode3.LeftChild
                    binaryNode3.LeftChild = leftChild1.RightChild
                    leftChild1.RightChild = binaryNode3
                    binaryNode3 = leftChild1
                    If binaryNode3.LeftChild Is Nothing Then
                        Exit While
                    End If
                End If

                binaryNode2.LeftChild = binaryNode3
                binaryNode2 = binaryNode3
                binaryNode3 = binaryNode3.LeftChild
                Continue While
            End If

            If tokenStart <= binaryNode3.Token.TokenStart OrElse binaryNode3.RightChild Is Nothing Then
                Exit While
            End If

            If tokenStart > binaryNode3.RightChild.Token.TokenStart Then
                Dim rightChild1 As BinaryNode = binaryNode3.RightChild
                binaryNode3.RightChild = rightChild1.LeftChild
                rightChild1.LeftChild = binaryNode3
                binaryNode3 = rightChild1
                If binaryNode3.RightChild Is Nothing Then
                    Exit While
                End If
            End If

            binaryNode1.RightChild = binaryNode3
            binaryNode1 = binaryNode3
            binaryNode3 = binaryNode3.RightChild
        End While

        binaryNode1.RightChild = binaryNode3.LeftChild
        binaryNode2.LeftChild = binaryNode3.RightChild
        binaryNode3.LeftChild = header.RightChild
        binaryNode3.RightChild = header.LeftChild
        Root = binaryNode3
    End Sub

    Friend Function UpdateKeys(editStart As Integer, insertCount As Integer) As Span
        Dim result As New Span(0, 0)
        If insertCount < 0 Then
            result = RemoveRange(editStart, editStart - insertCount)
            Dim num As Integer = result.Length + insertCount
            If num < 0 Then num = 0
            result = New Span(result.Start, num)
        End If

        If Root Is Nothing Then Return result

        If insertCount < 0 Then
            UpdateKeysRec(editStart - insertCount, insertCount)
        Else
            UpdateKeysRec(editStart, insertCount)
        End If
        Return result
    End Function

    Friend Sub UpdateKeysRec(thresHold As Integer, insertCount As Integer)
        If Root Is Nothing Then Return

        Dim stack1 As New Stack(Of BinaryNode)(count)
        stack1.Push(Root)

        While stack1.Count > 0
            Dim binaryNode1 As BinaryNode = stack1.Pop()
            If binaryNode1.RightChild IsNot Nothing Then
                stack1.Push(binaryNode1.RightChild)
            End If

            If binaryNode1.LeftChild IsNot Nothing AndAlso binaryNode1.Token.TokenStart >= thresHold Then
                stack1.Push(binaryNode1.LeftChild)
            End If
            binaryNode1.Token.ChangeTokenStart(thresHold, insertCount)
        End While

    End Sub

    Friend Function RemoveRange(minPosition As Integer, maxPosition As Integer) As Span
        If Root Is Nothing Then
            Return New Span(0, 0)
        End If

        Dim num As Integer = Integer.MaxValue
        Dim num2 As Integer = Integer.MinValue
        Dim list1 As New List(Of Integer)
        Dim stack1 As New Stack(Of BinaryNode)(count)
        stack1.Push(Root)
        Dim span1 As New Span(minPosition, maxPosition - minPosition)

        While stack1.Count > 0
            Dim binaryNode1 As BinaryNode = stack1.Pop()
            If binaryNode1.RightChild IsNot Nothing AndAlso binaryNode1.Token.TokenEnd <= maxPosition Then
                stack1.Push(binaryNode1.RightChild)
            End If

            If binaryNode1.LeftChild IsNot Nothing AndAlso binaryNode1.Token.TokenStart > minPosition Then
                stack1.Push(binaryNode1.LeftChild)
            End If

            If binaryNode1.Token.Span.OverlapsWith(span1) Then
                list1.Add(binaryNode1.Token.TokenStart)
                num = Math.Min(num, binaryNode1.Token.TokenStart)
                num2 = Math.Max(num2, binaryNode1.Token.TokenEnd)
            End If
        End While

        For Each item As Integer In list1
            Splay(item)
            If Root.LeftChild Is Nothing Then
                Root = Root.RightChild
                Continue For
            End If

            Dim rightChild1 As BinaryNode = Root.RightChild
            Root = Root.LeftChild
            Splay(maxPosition)
            _Root.RightChild = rightChild1
        Next

        count -= list1.Count
        If num = Integer.MaxValue Then num = 0
        If num2 = Integer.MinValue Then num2 = 0
        Return New Span(num, num2 - num)
    End Function

    Friend Function GetPrevTokenForPosition(position As Integer) As Token
        If Root Is Nothing Then Return Nothing

        Splay(position)
        If Root.Token.TokenStart <= position Then
            Return Root.Token
        End If

        If Root.LeftChild Is Nothing Then
            Return Nothing
        End If

        Dim binaryNode1 As BinaryNode = Root.LeftChild
        While binaryNode1.RightChild IsNot Nothing
            binaryNode1 = binaryNode1.RightChild
        End While

        Return binaryNode1.Token
    End Function

    Friend Function IsTokenEndAfter(position As Integer) As Boolean
        If Root Is Nothing Then
            Return False
        End If

        Splay(position)
        If Root.Token.TokenEnd > position Then
            Return True
        End If

        If Root.RightChild Is Nothing Then
            Return False
        End If

        Dim rightChild1 As BinaryNode = Root.RightChild
        While rightChild1.RightChild IsNot Nothing AndAlso rightChild1.Token.TokenEnd <= position
            rightChild1 = rightChild1.RightChild
        End While
        Return rightChild1.Token.TokenEnd > position
    End Function
End Class
