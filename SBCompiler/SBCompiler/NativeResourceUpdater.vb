Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.SmallVisualBasic

Public Class NativeResourceUpdater

    <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function BeginUpdateResource(fileName As String, bDeleteExistingResources As Boolean) As IntPtr
    End Function

    <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function UpdateResource(hUpdate As IntPtr, lpType As IntPtr, lpName As IntPtr, wLanguage As UShort, lpData As Byte(), cbData As UInteger) As Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function EndUpdateResource(hUpdate As IntPtr, fDiscard As Boolean) As Boolean
    End Function

    Public Shared Sub UpdateVersionResource(exePath As String, info As ProgramInfo)
        Dim hUpdate = BeginUpdateResource(exePath, False)
        If hUpdate = IntPtr.Zero Then
            Throw New System.ComponentModel.Win32Exception()
        End If

        Dim versionResource = BuildVersionResource(info)
        ' RT_VERSION is resource type 16.
        Dim RT_VERSION As New IntPtr(&H10)
        ' Resource name is usually 1.
        Dim RES_NAME As New IntPtr(1)
        ' Use nutral language.
        Dim language As UShort = 0

        If Not UpdateResource(hUpdate, RT_VERSION, RES_NAME, language, versionResource, CUInt(versionResource.Length)) Then
            Throw New System.ComponentModel.Win32Exception()
        End If

        If Not EndUpdateResource(hUpdate, False) Then
            Throw New System.ComponentModel.Win32Exception()
        End If
    End Sub

    Private Shared Function BuildVersionResource(info As ProgramInfo) As Byte()
        ' Build a minimal VS_VERSION_INFO resource that includes the fixed file info,
        ' one StringFileInfo block (with a single StringTable), and a VarFileInfo block.
        Using ms As New MemoryStream(), bw As New BinaryWriter(ms)
            Dim startPos As Long = ms.Position
            bw.Write(CUShort(0))      ' wLength (placeholder)
            bw.Write(CUShort(52))     ' wValueLength = size of VS_FIXEDFILEINFO (52 bytes)
            bw.Write(CUShort(0))      ' wType = binary
            WriteUnicodeString(bw, "VS_VERSION_INFO")
            AlignToDword(bw)

            ' --- VS_FIXEDFILEINFO (13 DWORDS: 52 bytes) ---
            bw.Write(&HFEEF04BDUI)    ' dwSignature
            bw.Write(&H10000UI)       ' dwStrucVersion

            Dim versionParts = info.Version.Split("."c)
            Dim major = If(versionParts.Length > 0, Convert.ToUInt16(versionParts(0)), 0)
            Dim minor = If(versionParts.Length > 1, Convert.ToUInt16(versionParts(1)), 0)
            Dim buildPart = If(versionParts.Length > 2, Convert.ToUInt16(versionParts(2)), 0)
            Dim revision = If(versionParts.Length > 3, Convert.ToUInt16(versionParts(3)), 0)

            Dim fileVersionMS = CUInt((major << 16) Or minor)
            Dim fileVersionLS = CUInt((buildPart << 16) Or revision)
            bw.Write(fileVersionMS)    ' dwFileVersionMS
            bw.Write(fileVersionLS)    ' dwFileVersionLS
            bw.Write(fileVersionMS)    ' dwProductVersionMS
            bw.Write(fileVersionLS)    ' dwProductVersionLS

            bw.Write(&H3FUI)           ' dwFileFlagsMask
            bw.Write(0UI)              ' dwFileFlags
            bw.Write(&H40004UI)        ' dwFileOS (VOS_NT_WINDOWS32)
            bw.Write(1UI)              ' dwFileType (VFT_APP)
            bw.Write(0UI)              ' dwFileSubtype
            bw.Write(0UI)              ' dwFileDateMS
            bw.Write(0UI)              ' dwFileDateLS

            ' --- StringFileInfo block ---
            Dim sfiBytes As Byte() = BuildStringFileInfo(info)
            bw.Write(sfiBytes)

            ' --- VarFileInfo block ---
            Dim vfiBytes As Byte() = BuildVarFileInfo()
            bw.Write(vfiBytes)

            ' Update the top-level wLength
            Dim totalLength As UShort = CUShort(ms.Length)
            ms.Position = startPos
            bw.Write(totalLength)

            Return ms.ToArray()
        End Using
    End Function

    Private Shared Sub WriteUnicodeString(bw As BinaryWriter, s As String)
        Dim bytes() As Byte = System.Text.Encoding.Unicode.GetBytes(s)
        bw.Write(bytes)
        bw.Write(CUShort(0)) ' null terminator
    End Sub

    Private Shared Sub AlignToDword(bw As BinaryWriter)
        Dim pos As Long = bw.BaseStream.Position
        Dim padding As Integer = CInt((4 - (pos Mod 4)) Mod 4)
        For i As Integer = 1 To padding
            bw.Write(CByte(0))
        Next
    End Sub

    Private Shared Function BuildStringFileInfo(info As ProgramInfo) As Byte()
        Using ms As New MemoryStream(), bw As New BinaryWriter(ms)
            Dim startPos As Long = ms.Position
            bw.Write(CUShort(0))   ' wLength (placeholder)
            bw.Write(CUShort(0))   ' wValueLength = 0 (container)
            bw.Write(CUShort(1))   ' wType = text
            WriteUnicodeString(bw, "StringFileInfo")
            AlignToDword(bw)

            Dim stBytes As Byte() = BuildStringTable(info)
            bw.Write(stBytes)

            Dim lengthValue As UShort = CUShort(ms.Length)
            ms.Position = startPos
            bw.Write(lengthValue)
            Return ms.ToArray()
        End Using
    End Function

    Private Shared Function BuildStringTable(info As ProgramInfo) As Byte()
        Using ms As New MemoryStream(), bw As New BinaryWriter(ms)
            Dim startPos As Long = ms.Position
            bw.Write(CUShort(0))   ' wLength (placeholder)
            bw.Write(CUShort(0))   ' wValueLength = 0 (container)
            bw.Write(CUShort(1))   ' wType = text
            WriteUnicodeString(bw, "000004B0") ' Language and codepage identifier.
            AlignToDword(bw)

            ' Add string entries.
            Dim entries As New List(Of Byte())
            entries.Add(BuildStringEntry("FileDescription", info.Description))
            entries.Add(BuildStringEntry("ProductName", info.Product))
            entries.Add(BuildStringEntry("ProductVersion", info.Version))
            entries.Add(BuildStringEntry("CompanyName", info.Company))
            entries.Add(BuildStringEntry("LegalCopyright", info.Copyright))
            entries.Add(BuildStringEntry("InternalName", info.Title))

            For Each entry As Byte() In entries
                bw.Write(entry)
                AlignToDword(bw)
            Next

            Dim lengthValue As UShort = CUShort(ms.Length)
            ms.Position = startPos
            bw.Write(lengthValue)
            Return ms.ToArray()
        End Using
    End Function

    Private Shared Function BuildStringEntry(key As String, value As String) As Byte()
        Using ms As New MemoryStream(), bw As New BinaryWriter(ms)
            Dim startPos As Long = ms.Position
            bw.Write(CUShort(0))            ' wLength (placeholder)
            bw.Write(CUShort(value.Length))   ' wValueLength = number of characters in value
            bw.Write(CUShort(1))            ' wType = text
            WriteUnicodeString(bw, key)
            AlignToDword(bw)
            WriteUnicodeString(bw, value)
            AlignToDword(bw)

            Dim lengthValue As UShort = CUShort(ms.Length)
            ms.Position = startPos
            bw.Write(lengthValue)
            Return ms.ToArray()
        End Using
    End Function

    Private Shared Function BuildVarFileInfo() As Byte()
        Using ms As New MemoryStream(), bw As New BinaryWriter(ms)
            Dim startPos As Long = ms.Position
            bw.Write(CUShort(0)) ' wLength (placeholder)
            bw.Write(CUShort(0)) ' wValueLength = 0 (container)
            bw.Write(CUShort(1)) ' wType = text
            WriteUnicodeString(bw, "VarFileInfo")
            AlignToDword(bw)

            Dim transBytes As Byte() = BuildTranslationBlock()
            bw.Write(transBytes)

            Dim lengthValue As UShort = CUShort(ms.Length)
            ms.Position = startPos
            bw.Write(lengthValue)
            Return ms.ToArray()
        End Using
    End Function

    Private Shared Function BuildTranslationBlock() As Byte()
        Using ms As New MemoryStream(), bw As New BinaryWriter(ms)
            Dim startPos As Long = ms.Position
            bw.Write(CUShort(0)) ' wLength (placeholder)
            bw.Write(CUShort(4)) ' wValueLength = 4 bytes (two WORDs)
            bw.Write(CUShort(0)) ' wType = binary
            WriteUnicodeString(bw, "Translation")
            AlignToDword(bw)
            bw.Write(CUShort(&H0)) ' Language: nutral
            bw.Write(CUShort(&H4B0)) ' Code page: 1252
            AlignToDword(bw)
            Dim lengthValue As UShort = CUShort(ms.Length)
            ms.Position = startPos
            bw.Write(lengthValue)
            Return ms.ToArray()
        End Using
    End Function

End Class