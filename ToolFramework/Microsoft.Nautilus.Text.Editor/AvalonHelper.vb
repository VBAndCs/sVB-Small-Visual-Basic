Imports System.Runtime.InteropServices
Imports System.Windows
Imports System.Windows.Interop
Imports System.Windows.Media
Imports System.Windows.Threading

Namespace Microsoft.Nautilus.Text.Editor

    Public NotInheritable Class AvalonHelper
        Private Shared _threadMgr As NativeMethods.ITfThreadMgr
        Private Shared _threadMgrFailed As Boolean
        Private Shared _documentMgr As IntPtr

        Public Shared Function GetRootVisual(visual As Visual) As Visual
            If visual Is Nothing Then
                Throw New ArgumentNullException("visual")
            End If

            Dim parent As DependencyObject = visual
            Dim result As Visual = visual
            Do
                parent = VisualTreeHelper.GetParent(parent)
                If parent Is Nothing Then Exit Do

                If TypeOf parent Is Visual Then
                    result = CType(parent, Visual)
                End If
            Loop

            Return result
        End Function

        Public Shared Function GetHwndSource(visual As Visual) As HwndSource
            If visual Is Nothing Then
                Throw New ArgumentNullException("visual")
            End If

            Dim rootVisual As Visual = GetRootVisual(visual)
            Return CType(PresentationSource.FromVisual(rootVisual), HwndSource)
        End Function

        Public Shared Function GetImmContext(visual As Visual) As IntPtr
            If visual Is Nothing Then
                Throw New ArgumentNullException("visual")
            End If

            Dim hwndSource1 As HwndSource = GetHwndSource(visual)
            If hwndSource1 IsNot Nothing Then
                Return NativeMethods.ImmGetContext(hwndSource1.Handle)
            End If
            Return IntPtr.Zero
        End Function

        Public Shared Function ReleaseImmContext(visual As Visual, immContext As IntPtr) As Boolean
            If visual Is Nothing Then
                Throw New ArgumentNullException("visual")
            End If

            If immContext = IntPtr.Zero Then
                Return False
            End If

            Dim hwndSource1 As HwndSource = GetHwndSource(visual)
            If hwndSource1 IsNot Nothing Then
                Return NativeMethods.ImmReleaseContext(hwndSource1.Handle, immContext)
            End If
            Return False
        End Function

        Public Shared Function SetImmFontHeight(immContext As IntPtr, fontHeight As Integer) As Boolean
            If immContext = IntPtr.Zero Then
                Return False
            End If

            Dim lOGFONT1 As New NativeMethods.LOGFONT
            lOGFONT1.lfHeight = fontHeight
            Dim intPtr1 As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(CObj(lOGFONT1)))
            Marshal.StructureToPtr(CObj(lOGFONT1), intPtr1, fDeleteOld:=True)
            Dim result As Boolean = NativeMethods.ImmSetCompositionFontW(immContext, intPtr1)
            Marshal.FreeHGlobal(intPtr1)
            Return result
        End Function

        Public Shared Function SetCompositionWindowPosition(immContext As IntPtr, pt As Point, relativeTo As Visual) As Boolean
            If relativeTo Is Nothing Then
                Throw New ArgumentNullException("relativeTo")
            End If

            Dim rootVisual As Visual = GetRootVisual(relativeTo)
            If rootVisual Is Nothing Then
                Return False
            End If

            Dim point1 As Point = relativeTo.TransformToAncestor(rootVisual).Transform(pt)
            Dim cOMPOSITIONFORM1 As NativeMethods.COMPOSITIONFORM = CType(Nothing, NativeMethods.COMPOSITIONFORM)
            cOMPOSITIONFORM1.dwStyle = 2
            cOMPOSITIONFORM1.ptCurrentPos = CType(Nothing, NativeMethods.POINT)
            cOMPOSITIONFORM1.ptCurrentPos.x = CInt(CLng(Fix(point1.X)) Mod Integer.MaxValue)
            cOMPOSITIONFORM1.ptCurrentPos.y = CInt(CLng(Fix(point1.Y)) Mod Integer.MaxValue)
            Dim intPtr1 As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(CObj(cOMPOSITIONFORM1)))
            Marshal.StructureToPtr(CObj(cOMPOSITIONFORM1), intPtr1, fDeleteOld:=True)
            Dim result As Boolean = NativeMethods.ImmSetCompositionWindow(immContext, intPtr1)
            Marshal.FreeHGlobal(intPtr1)
            Return result
        End Function

        Public Shared Sub EnableImmComposition(dispatcher As DispatcherObject)
            If dispatcher Is Nothing Then
                Throw New ArgumentNullException("dispatcher")
            End If

            If _threadMgrFailed Then
                Return
            End If

            If _threadMgr Is Nothing Then
                NativeMethods.TF_CreateThreadMgr(_threadMgr)
                If _threadMgr Is Nothing Then
                    _threadMgrFailed = True
                    Return
                End If
            End If

            dispatcher.Dispatcher.BeginInvoke(DispatcherPriority.Background, CType(
                    Function() As Object
                        _threadMgr.GetFocus(_documentMgr)
                        _threadMgr.SetFocus(IntPtr.Zero)
                        Return Nothing
                    End Function, DispatcherOperationCallback), Nothing)
        End Sub

        Public Shared Sub DisableImmComposition(dispatcher As DispatcherObject)
            If dispatcher Is Nothing Then
                Throw New ArgumentNullException("dispatcher")
            End If

            If _threadMgrFailed Then Return

            If _threadMgr Is Nothing Then
                NativeMethods.TF_CreateThreadMgr(_threadMgr)
                If _threadMgr Is Nothing Then
                    _threadMgrFailed = True
                    Return
                End If
            End If

            dispatcher.Dispatcher.BeginInvoke(DispatcherPriority.Background, CType(
                        Function() As Object
                            Dim docMgr As IntPtr = Nothing
                            _threadMgr.GetFocus(docMgr)
                            If docMgr = IntPtr.Zero Then
                                _threadMgr.SetFocus(_documentMgr)
                            End If
                            Return Nothing
                        End Function, DispatcherOperationCallback), Nothing)
        End Sub

    End Class
End Namespace
