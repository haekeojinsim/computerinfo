Imports System.Net.NetworkInformation
Imports System.Management
Imports Word = Microsoft.Office.Interop.Word
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.Net
Imports System.Text.RegularExpressions
Imports Microsoft.Win32

Public Class Form1
    '===============================================================
    '=  ComputerInfo
    '=  Version 2.1
    '=  Copyright (C) 2017 Haekeo Jimsim
    '=  Date Compiled: 26/4/2017 at 3:35AM
    '===============================================================
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '------------------------------------------------------------------------------
        '  Display Outputs
        '------------------------------------------------------------------------------
        'Left-hand Part
        ShowDate() 'Display Date
        ShowComputerName() 'Display Computer Name
        ShowUserName() 'Display User Name
        ShowOS() 'Display Windows Operating System
        ShowServicePack() 'Display Windows Service Packs
        ShowBrandName() 'Display Brand Name
        ShowModel() 'Display Model
        ShowProcessor() 'Display Processor Name
        ShowRAM() 'Display RAM Size
        ShowSerialNumber() 'Display Serial Number
        ShowDefaultPrinter() 'Display Default Printer

        'Right-hand Part
        ShowIPAddress() 'Display IP Address
        ShowMACAddress() 'Get MAC Address
        ShowPrimaryDNSAddress() 'Display primary DNS Address
        ShowSecondaryDNSAddress() 'Display secondary DNS Address
        DisplayHardDiskSpace() 'Display Hard Disk Space
        ShowImageVersion() 'Display Image Versions
        ShowDriveCEncryptionStatus() 'Display Drive C encryption status
        ShowDriveDEncryptionStatus() 'Display Drive D encryption status
        ShowUSBState() 'Display the status of USB

        'Additional
        ExportHeader() 'Temporary export Header.PNG onto current directory
    End Sub

    '-------------------------------------------------------------------------------------------
    '  Display Functions
    '-------------------------------------------------------------------------------------------
    Function ShowDate() 'Display Date
        Try
            DateResult.Text = GetDate().ToString
        Catch ex As Exception
            Return Nothing
        End Try
        Return Nothing
    End Function

    Function ShowComputerName() 'Display Computer Name
        Try
            ComputerNameResult.Text = GetComputerName().ToString
        Catch ex As Exception
            Return Nothing
        End Try
        Return Nothing
    End Function

    Function ShowOS() 'Display Windows Operating System
        Try
            OSResult.Text = GetOperatingSystem().ToString
        Catch ex As Exception
            Return Nothing
        End Try
        Return Nothing
    End Function

    Function ShowServicePack() 'Display Windows Service Packs
        Try
            SPResult.Text = GetWindowsServicePack().ToString
        Catch ex As Exception
            Return Nothing
        End Try
        Return Nothing
    End Function

    Function ShowBrandName() 'Display Brand Name
        Try
            BrandNameResult.Text = GetBrandName().ToString
        Catch ex As Exception
            Return Nothing
        End Try
        Return Nothing
    End Function


    Function ShowModel() 'Display Model
        Try
            ModelResult.Text = GetModel().ToString
        Catch ex As Exception
            Return Nothing
        End Try
        Return Nothing
    End Function

    Function ShowProcessor() 'Display Processor Name
        Try
            ProcessorsResult.Text = GetProcessorName().ToString
        Catch ex As Exception
            Return Nothing
        End Try
        Return Nothing
    End Function

    Function ShowRAM() 'Display RAM Size
        Try
            RAMResult.Text = GetRAMSize().ToString
        Catch ex As Exception
            Return Nothing
        End Try
        Return Nothing
    End Function

    Function ShowSerialNumber() 'Display Serial Number
        Try
            SerialNumberResult.Text = GetSerialNumber().ToString
        Catch ex As Exception
            Return Nothing
        End Try
        Return Nothing
    End Function

    Function ShowDefaultPrinter() 'Display Default Printer
        Try
            DefaultPrinterResult.Text = GetDefaultPrinter().ToString
        Catch ex As Exception
            Return Nothing
        End Try
        Return Nothing
    End Function

    Function ShowUserName() 'Display User Name
        Try
            UserNameResult.Text = GetUserName().ToString
        Catch ex As Exception
            Return Nothing
        End Try
        Return Nothing
    End Function

    Function ShowIPAddress() 'Display IP Address
        Dim IPAddress As String = GetIPAddressIPv4.ToString
        Dim LocalIPAddress As String = "127.0.0.1"

        Try
            If IPAddress = LocalIPAddress Then
                IPAddressResult.Text = "Local only"
            Else
                IPAddressResult.Text = GetIPAddressIPv4().ToString
            End If
        Catch ex As Exception
            IPAddressResult.Text = "Not connected"
        End Try
        Return Nothing
    End Function

    Function ShowMACAddress() 'Get MAC Address
        Try
            MACResult.Text = GetMACAddress().ToString
        Catch ex As Exception
            MACResult.Text = "Adapter disabled"
        End Try
        Return Nothing
    End Function

    Function ShowPrimaryDNSAddress() 'Display primary DNS Address
        Try
            PrimaryDNSResult.Text = GetPrimaryDNSAddress().ToString
        Catch ex As Exception
            PrimaryDNSResult.Text = "Not available"
        End Try
        Return Nothing
    End Function

    Function ShowSecondaryDNSAddress() 'Display secondary DNS Address
        Try
            SecondaryDNSResult.Text = GetSecondaryDNSAddress().ToString
        Catch ex As Exception
            SecondaryDNSResult.Text = "Not available"
        End Try
        Return Nothing
    End Function

    Function ShowImageVersion() 'Display Image Versions
        Try
            ImageVersionResult.Text = GetImageVersion().ToString
        Catch ex As Exception
            Return Nothing
        End Try
        Return Nothing
    End Function

    Function ShowDriveCEncryptionStatus() 'Display Drive C encryption status
        Try
            DriveCEncryptionResult.Text = GetDriveCEncryptionStatus().ToString
        Catch ex As Exception
            Return Nothing
        End Try
        Return Nothing
    End Function

    Function ShowDriveDEncryptionStatus() 'Display Drive D encryption status
        Try
            DriveDEncryptionResult.Text = GetDriveDEncryptionStatus().ToString
        Catch ex As Exception
            Return Nothing
        End Try
        Return Nothing
    End Function

    Function ShowUSBState() 'Display USB status
        Try
            USBStatusResult.Text = GetUSBState().ToString
        Catch ex As Exception
            Return Nothing
        End Try
        Return Nothing
    End Function

    '-------------------------------------------------------------------------------------------
    '  Core Getter Functions
    '-------------------------------------------------------------------------------------------
    Function GetDate() 'Get Date
        Dim regDate As Date = Date.Now()
        Dim SavedDate As String
        Dim strDate As String = regDate.ToString("dd\/MM\/yyyy")
        SavedDate = strDate 'Save date into variables
        Return SavedDate
    End Function

    Function GetComputerName() 'Get Computer Name
        Dim SavedComputerName As String
        SavedComputerName = My.Computer.Name 'Save Computer Name into variables
        Return SavedComputerName
    End Function

    Function GetOperatingSystem() 'Get Operating System
        Dim SavedOSName As String
        SavedOSName = My.Computer.Info.OSFullName 'Save Windows Operating System into variables
        Return SavedOSName
    End Function
    Function GetWindowsServicePack() 'Get Windows Service Pack
        Dim ServicePack As String
        Dim SavedServicePack As String
        Dim CurrentVersion As String
        Dim ReleaseID As String
        Dim CurrentBuild As String
        ServicePack = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CSDVersion", Nothing)
        CurrentVersion = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentVersion", Nothing)
        ReleaseID = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ReleaseId", Nothing)
        CurrentBuild = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentBuild", Nothing)
        SavedServicePack = ServicePack 'Save the Service Pack into the variables

        If SavedServicePack Is Nothing Then
            ServicePack = CurrentVersion + " (Version " + ReleaseID + " Build " + CurrentBuild + ")"
            SavedServicePack = ServicePack
        Else
            ServicePack = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CSDVersion", Nothing)
            SavedServicePack = ServicePack
        End If

        Return SavedServicePack
    End Function

    Function GetBrandName() 'Get Brand Name
        Dim objCS As Management.ManagementObjectSearcher
        Dim manufacturerName As String
        Dim SavedManufacturerName As String
        objCS = New Management.ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem")
        For Each objMgmt In objCS.Get
            manufacturerName = objMgmt("manufacturer").ToString()
            SavedManufacturerName = manufacturerName 'Save the Manufacturer Name into the variables
            Return manufacturerName
        Next
        Return Nothing
    End Function

    Function GetModel() 'Get Model
        Dim objCS As Management.ManagementObjectSearcher
        Dim model As String
        Dim SavedModel As String
        objCS = New Management.ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem")
        For Each objMgmt In objCS.Get
            model = objMgmt("model").ToString()
            SavedModel = model 'Save the Model into the variables
            Return model
        Next
        Return Nothing
    End Function

    Function GetProcessorName() 'Get Processor Name
        Dim ProcessorName As String
        Dim SavedProcessorName As String
        ProcessorName = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\SYSTEM\CentralProcessor\0", "ProcessorNameString", Nothing)
        SavedProcessorName = ProcessorName 'Save the Processor Name into the variables
        Return SavedProcessorName
    End Function

    Function GetRAMSize() 'Get RAM Size
        Dim RAMSize As String
        Dim SavedRAMSize As String
        RAMSize = setPrefix(My.Computer.Info.TotalPhysicalMemory)
        SavedRAMSize = RAMSize 'Save RAM Size into variables
        Return RAMSize
    End Function

    Function GetSerialNumber() 'Get Serial Number

        Dim q As New SelectQuery("Win32_bios")
        Dim search As New ManagementObjectSearcher(q)
        Dim SerialNumber As New ManagementObject
        Dim SavedSerialNumber As String
        For Each SerialNumber In search.Get
            SavedSerialNumber = SerialNumber("serialnumber").ToString 'Save serial number into variables
            Return SavedSerialNumber
        Next
        Return Nothing
    End Function

    Function GetDefaultPrinterFunction() 'Get Default Printer (First Step)
        Dim oPS As New System.Drawing.Printing.PrinterSettings
        Dim NoPrinterName = "No default printers"
        Try
            Return oPS.PrinterName
        Catch ex As System.Exception
            Return NoPrinterName
        Finally
            oPS = Nothing
        End Try
    End Function

    Function GetDefaultPrinter() 'Get Default Printer (Second Step)
        Dim DefaultPrinterName As String
        Dim SavedPrinterName As String
        DefaultPrinterName = GetDefaultPrinterFunction()
        SavedPrinterName = DefaultPrinterName 'Save Default Printer Name into variables
        Return SavedPrinterName
    End Function

    Function GetUserName() 'Get User Name
        Dim SavedUserName As String
        SavedUserName = SystemInformation.UserName 'Save User Name into variables
        Return SavedUserName
    End Function

    Function GetIPAddressIPv4() 'Get IP Address Version 4
        Dim IPAddress As String = String.Empty
        Dim strHostName As String = System.Net.Dns.GetHostName()
        Dim iphe As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(strHostName)

        For Each ipheal As System.Net.IPAddress In iphe.AddressList
            If ipheal.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                IPAddress = ipheal.ToString()
            End If
        Next
        Return IPAddress
    End Function

    Function GetMACAddressxx() 'Get MAC Address
        Dim MACAddress As String = " "
        Dim NetworkCard As NetworkInterface
        Dim GatewayAddress As GatewayIPAddressInformation

        For Each NetworkCard In NetworkInterface.GetAllNetworkInterfaces
            For Each GatewayAddress In NetworkCard.GetIPProperties.GatewayAddresses
                If GatewayAddress.Address.ToString <> "0.0.0.0" And NetworkCard.OperationalStatus.ToString() = "Up" Then
                    MACAddress = NetworkCard.GetPhysicalAddress.ToString
                    Return MACAddress

                Else
                    Return "Not connected"
                End If
            Next
        Next
        Return Nothing
    End Function

    Function GetMACAddress()
        Dim mc As New ManagementClass("Win32_NetworkAdapterConfiguration")
        Dim moc As ManagementObjectCollection = mc.GetInstances()
        Dim MACAddress As String = [String].Empty
        For Each mo As ManagementObject In moc
            If MACAddress = [String].Empty Then
                ' only return MAC Address from first card
                If CBool(mo("IPEnabled")) = True Then
                    MACAddress = mo("MacAddress").ToString()
                End If
            End If
            mo.Dispose()
        Next

        Return MACAddress
    End Function

    Function GetPrimaryDNSAddress() 'Get primary DNS Address
        Dim wmi As ManagementClass
        Dim obj As ManagementObject
        Dim objs As ManagementObjectCollection
        Dim PrimaryDNS As String

        wmi = New ManagementClass("Win32_NetworkAdapterConfiguration")
        objs = wmi.GetInstances()
        For Each obj In objs
            If Not IsNothing(obj("DNSServerSearchOrder")) AndAlso UBound(obj("DNSServerSearchOrder")) >= 0 Then
                PrimaryDNS = "" & obj("DNSServerSearchOrder")(0).ToString
                Return PrimaryDNS
            End If
        Next
        objs.Dispose()
        wmi.Dispose()
        Return Nothing
    End Function

    Function GetSecondaryDNSAddress() 'Get secondary DNS Address
        Dim wmi As ManagementClass
        Dim obj As ManagementObject
        Dim objs As ManagementObjectCollection
        Dim SecondaryDNS As String

        wmi = New ManagementClass("Win32_NetworkAdapterConfiguration")
        objs = wmi.GetInstances()
        For Each obj In objs
            If Not IsNothing(obj("DNSServerSearchOrder")) AndAlso UBound(obj("DNSServerSearchOrder")) >= 1 Then
                SecondaryDNS = "" & obj("DNSServerSearchOrder")(1).ToString
                Return SecondaryDNS
            End If
        Next
        objs.Dispose()
        wmi.Dispose()
        Return Nothing
    End Function

    Function GetDriveCTotalSpace() 'Get Drive C Total Size
        Dim cdrive As System.IO.DriveInfo
        cdrive = My.Computer.FileSystem.GetDriveInfo("C:\")

        Try
            Return cdrive.TotalSize.ToString()
        Catch ex As Exception
            Return "Not detected"
        End Try
    End Function

    Function GetDriveDTotalSpace() 'Get Drive D Total Size
        Dim ddrive As System.IO.DriveInfo
        ddrive = My.Computer.FileSystem.GetDriveInfo("D:\")

        Try
            Return ddrive.TotalSize.ToString()
        Catch ex As Exception
            Return "Not detected"
        End Try
    End Function

    Function GetDriveCFreeSpace() 'Get Drive C Free Space
        Dim cdrive As System.IO.DriveInfo
        cdrive = My.Computer.FileSystem.GetDriveInfo("C:\")

        Try
            Return cdrive.TotalFreeSpace.ToString()
        Catch ex As Exception
            Return "Not detected"
        End Try
    End Function

    Function GetDriveDFreeSpace() 'Get Drive D Free Space
        Dim ddrive As System.IO.DriveInfo
        ddrive = My.Computer.FileSystem.GetDriveInfo("D:\")

        Try
            Return ddrive.TotalFreeSpace.ToString()
        Catch ex As Exception
            Return "Not detected"
        End Try
    End Function

    Function DisplayHardDiskSpace() 'Display hard disk space
        Dim DriveCTotal As String = GetDriveCTotalSpace()
        Dim DriveDTotal As String = GetDriveDTotalSpace()
        Dim DriveCFreeSpace As String = GetDriveCFreeSpace()
        Dim DriveDFreeSpace As String = GetDriveDFreeSpace()
        Dim DriveNotDetected As String = "Not detected"

        If DriveCTotal = DriveNotDetected Or DriveCFreeSpace = DriveNotDetected Or DriveDTotal = DriveNotDetected Or DriveDFreeSpace = DriveNotDetected Then
            DriveCResult.Text = "Not detected"
            DriveDResult.Text = "Not detected"
        Else
            DriveCTotal = setPrefix(DriveCTotal)
            DriveCFreeSpace = setPrefix(DriveCFreeSpace)
            DriveCResult.Text = DriveCFreeSpace + " out of " + DriveCTotal

            DriveDTotal = setPrefix(DriveDTotal)
            DriveDFreeSpace = setPrefix(DriveDFreeSpace)
            DriveDResult.Text = DriveDFreeSpace + " out of " + DriveDTotal
        End If

        Return Nothing
    End Function

    Function ExportDriveC() 'Export drive C space
        Dim DriveCTotal As String = GetDriveCTotalSpace()
        Dim DriveCFreeSpace As String = GetDriveCFreeSpace()
        Dim DriveCString As String
        Dim DriveNotDetected As String = "Not detected"

        If DriveCTotal = DriveNotDetected Or DriveCFreeSpace = DriveNotDetected Then
            DriveCString = "Not detected"
        Else
            DriveCTotal = setPrefix(DriveCTotal)
            DriveCFreeSpace = setPrefix(DriveCFreeSpace)
            DriveCString = DriveCFreeSpace + " out of " + DriveCTotal
        End If

        Return DriveCString
    End Function

    Function ExportDriveD() 'Export drive D space
        Dim DriveDTotal As String = GetDriveDTotalSpace()
        Dim DriveDFreeSpace As String = GetDriveDFreeSpace()
        Dim DriveDString As String
        Dim DriveNotDetected As String = "Not detected"

        If DriveDTotal = DriveNotDetected Or DriveDFreeSpace = DriveNotDetected Then
            DriveDString = "Not detected"
        Else
            DriveDTotal = setPrefix(DriveDTotal)
            DriveDFreeSpace = setPrefix(DriveDFreeSpace)
            DriveDString = DriveDFreeSpace + " out of " + DriveDTotal
        End If

        Return DriveDString
    End Function

    Function GetImageVersion() 'Retrieve Image Base Release from Windows Registry
        Dim ImageVersion As String
        Dim SavedImageName As String
        Dim Not As String = "Not found"
        ImageVersion = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation", "Model", Nothing) 'Get the Registry Key Values

        If ImageVersion Is Nothing Then 'If there is no Base Release value
            Return Not
        Else
            SavedImageName = ImageVersion 'Save the image version into the variables
        End If
        Return SavedImageName
    End Function

    Function GetDriveCEncryptionStatus() 'Retrieve Drive C encryption status from Windows Registry
        Dim EncryptionStatus As String = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\McAfee EndPoint Encryption\MfeEpePC\Status", "CryptState", Nothing) 'Get the Registry Key Values
        Dim ActivationStatus As String = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\McAfee EndPoint Encryption\MfeEpePC\Status", "Activated", Nothing)

        If ActivationStatus.Contains("Yes") Then
            If EncryptionStatus.Contains("Volume=C:,State=Encrypted") Then
                Return "Encrypted"
            Else
                Return "Decrypted"
            End If
        Else
            Return "Check McAfee status"
        End If
        Return Nothing
    End Function

    Function GetDriveDEncryptionStatus() 'Retrieve Drive D encryption status from Windows Registry
        Dim EncryptionStatus As String = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\McAfee EndPoint Encryption\MfeEpePc\Status", "CryptState", Nothing) 'Get the Registry Key Values
        Dim ActivationStatus As String = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\McAfee EndPoint Encryption\MfeEpePc\Status", "Activated", Nothing)

        If ActivationStatus.Contains("Yes") Then
            If EncryptionStatus.Contains("Volume=D:,State=Encrypted") Then
                Return "Encrypted"
            Else
                Return "Decrypted"
            End If
        Else
            Return "Check McAfee status"
        End If
        Return Nothing
    End Function

    Function GetUSBState() 'Retrieve USB status from Windows Registry
        Dim EncryptionStatus As String = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\USBSTOR", "Start", Nothing) 'Get the Registry Key Values
        Dim USBEnabled As String = "3"
        Dim USBDisabled As String = "4"

        If EncryptionStatus = USBDisabled Then
            Return "0x4"
        ElseIf EncryptionStatus = USBEnabled Then
            Return "0x3"
        Else
            Return "Unknown"
        End If
        Return Nothing
    End Function

    Function GetAppList() 'Retrieve list of installed apps from Windows Registry
        Dim regkey, subkey As RegistryKey
        Dim value As String
        Dim regpath As String = "Software\Microsoft\Windows\CurrentVersion\Uninstall"
        Dim includes As Boolean
        regkey = My.Computer.Registry.LocalMachine.OpenSubKey(regpath)
        Dim subkeys() As String = regkey.GetSubKeyNames

        For Each subk As String In subkeys
            subkey = regkey.OpenSubKey(subk)
            value = subkey.GetValue("DisplayName", "")

            If value <> "" Then
                includes = True
                If value.IndexOf("Hotfix") <> -1 Then includes = False
                If value.IndexOf("Security Update") <> -1 Then includes = False
                If value.IndexOf("Update for") <> -1 Then includes = False
                If value.IndexOf("Cumulative Security Update") <> -1 Then includes = False
                If includes = True Then Return value
            End If
        Next
        Return Nothing
    End Function

    Function ExportAppList() 'Export installed app onto text file
        Dim SavePath As String = Application.ExecutablePath & GetComputerName() + ".txt"
        Dim StreamWriter As New IO.StreamWriter(SavePath)
        Dim regkey, subkey As RegistryKey
        Dim value As String
        Dim regpath As String = "Software\Microsoft\Windows\CurrentVersion\Uninstall"
        Dim includes As Boolean
        regkey = My.Computer.Registry.LocalMachine.OpenSubKey(regpath)
        Dim subkeys() As String = regkey.GetSubKeyNames

        For Each subk As String In subkeys
            subkey = regkey.OpenSubKey(subk)
            value = subkey.GetValue("DisplayName", "")

            If value <> "" Then
                includes = True
                If value.IndexOf("Hotfix") <> -1 Then includes = False
                If value.IndexOf("Security Update") <> -1 Then includes = False
                If value.IndexOf("Update for") <> -1 Then includes = False
                If value.IndexOf("Cumulative Security Update") <> -1 Then includes = False
                If includes = True Then
                    StreamWriter.WriteLine(value)
                End If
            End If
        Next

        'Cleanup after writing result onto a text file
        StreamWriter.Close()
        StreamWriter.Dispose()
        Return Nothing
    End Function

    Function SendEmail() 'Export result onto email
        Dim OfficeType As Type
        OfficeType = Type.GetTypeFromProgID("Outlook.Application") 'Get the Outlook properties

        If OfficeType Is Nothing Then 'Check if Outlook is installed
            MessageBox.Show("Please install Microsoft Outlook and then try again.", “ComputerInfo")
        Else
            'Check if the text field is empty
            If (EngineerNameInput.Text.Trim = String.Empty) Or (EmployeeNameInput.Text.Trim = String.Empty) Or (EmployeeIDInput.Text.Trim = String.Empty) Or (ContactNumberInput.Text.Trim = String.Empty) Or (LANIDInput.Text.Trim = String.Empty) Or (DivisionInput.Text.Trim = String.Empty) Or (DepartmentInput.Text.Trim = String.Empty) Or (LocationInput.Text.Trim = String.Empty) Or (CostCentreInput01.Text.Trim = String.Empty) Or
           (String.IsNullOrEmpty(EngineerNameInput.Text.Trim)) Or (String.IsNullOrEmpty(EmployeeNameInput.Text.Trim)) Or (String.IsNullOrEmpty(EmployeeIDInput.Text.Trim)) Or (String.IsNullOrEmpty(ContactNumberInput.Text.Trim)) Or (String.IsNullOrEmpty(LANIDInput.Text.Trim)) Or (String.IsNullOrEmpty(DivisionInput.Text.Trim)) Or (String.IsNullOrEmpty(DepartmentInput.Text.Trim)) Or (String.IsNullOrEmpty(LocationInput.Text.Trim)) Or (String.IsNullOrEmpty(CostCentreInput01.Text.Trim)) Then
                MessageBox.Show("Please fill in all the required details..", "ComputerInfo")

            Else
                Dim Outl As Object
                Dim oInspector As Outlook.Inspector
                Outl = CreateObject("Outlook.Application")

                If Outl IsNot Nothing Then
                    Dim omsg As Object
                    omsg = Outl.CreateItem(0) '=Outlook.OlItemType.olMailItem'
                    oInspector = omsg.GetInspector
                    Try
                        omsg.Display(False) 'Does not lock Outlook (True locks to new email message)
                        'Set message properties here...'
                        omsg.To = "example@example.com"
                        omsg.subject = "Computer Inventory Checklist For " + GetComputerName() + ""
                        omsg.body = "Hi IT," & vbNewLine & "" & vbNewLine & "Attached Is the computer inventory checklist For " + GetComputerName() + "." + vbNewLine & vbNewLine &
                    "Thanks With regards," & vbNewLine &
                    GetComputerName() & vbNewLine & "" & vbNewLine &
                    "This message was automatically generated by ComputerInfo. Please don't reply to this message."

                        ' Add an attachment
                        'Replace with a valid attachment path.
                        Dim sSource As String = Application.ExecutablePath & GetComputerName() + ".pdf"
                        'Replace with attachment name
                        Dim sDisplayName As String = Application.ExecutablePath & GetComputerName() + ".pdf"

                        Dim sBodyLen As String = omsg.Body.Length
                        Dim oAttachs As Outlook.Attachments = omsg.Attachments
                        Dim oAttach As Outlook.Attachment
                        oAttach = oAttachs.Add(sSource, , sBodyLen + 1, sDisplayName)

                        omsg.Display(True) 'Display message to user

                    Catch COMex As System.Runtime.InteropServices.COMException
                        Return Nothing
                    Catch ex As Exception
                        Return Nothing
                    End Try
                End If
            End If
        End If
        Return Nothing
    End Function

    Function StopWord() 'Terminate Microsoft Word process
        Dim pList() As System.Diagnostics.Process =
    System.Diagnostics.Process.GetProcessesByName("winword")
        For Each proc As System.Diagnostics.Process In pList
            proc.Kill()
        Next
        Return Nothing
    End Function

    Function StopNotepad() 'Terminate Notepad process
        Dim pList() As System.Diagnostics.Process =
    System.Diagnostics.Process.GetProcessesByName("notepad")
        For Each proc As System.Diagnostics.Process In pList
            proc.Kill()
        Next
        Return Nothing
    End Function

    Function ExportHeader() 'Export Header.png to current directory
        My.Resources.ICHeader.Save("Header.png", Imaging.ImageFormat.Png)
        Return Nothing
    End Function

    Function DeleteICHeader() 'Delete Header.png in current directory
        Dim FileToDelete As String

        FileToDelete = Application.StartupPath & "\ICHeader.png" 'Image path

        If System.IO.File.Exists(FileToDelete) = True Then 'Check if file existsed
            System.IO.File.Delete(FileToDelete)
        End If
        Return Nothing
    End Function

    '------------------------------------------------------------------------------------------
    '  Get Input Field Value
    '------------------------------------------------------------------------------------------
    Function GetEmployeeName() 'Get Employee Name
        Dim EmployeeName As String
        EmployeeName = EmployeeNameInput.Text
        Return EmployeeName
    End Function

    Function GetEmployeeID() 'Get Employee ID
        Dim EmployeeID As String
        EmployeeID = EmployeeIDInput.Text
        Return EmployeeID
    End Function

    Function GetContactNumber() 'Get Contact Number
        Dim ContactNumber As String
        ContactNumber = ContactNumberInput.Text
        Return ContactNumber
    End Function

    Function GetLANID() 'Get LAN ID
        Dim LANID As String
        LANID = LANIDInput.Text
        Return LANID
    End Function

    Function GetDivision() 'Get Division
        Dim Division As String
        Division = DivisionInput.Text
        Return Division
    End Function

    Function GetDepartment() 'Get Department
        Dim Department As String
        Department = DepartmentInput.Text
        Return Department
    End Function

    Function GetLocation() 'Get Location
        Dim Location As String
        Location = LocationInput.Text
        Return Location
    End Function

    Function GetCallCentre() 'Get Cost Centre
        Dim CallCentre As String
        CallCentre = CostCentreInput01.Text + "-" + CostCentreInput02.Text + "-" + CostCentreInput03.Text
        Return CallCentre
    End Function
    Function GetEngineerName() 'Get Engineer Name
        Dim EngineerName As String
        EngineerName = EngineerNameInput.Text
        Return EngineerName
    End Function

    Function GetChecklistType() 'Get checklist type
        Dim ChecklistType As String

        If NewComputerRadio.Checked Then
            ChecklistType = "New Computer"

        ElseIf ExistingComputerRadio.Checked Then
            ChecklistType = "Existing Computer"

        Else
            ChecklistType = "Not Selected"

        End If
        Return ChecklistType
    End Function

    '------------------------------------------------------------------------------------------
    '  Others
    '------------------------------------------------------------------------------------------
    'Used to convert and set the prefix for the hard drive size
    'Sources: http://forum.codecall.net/topic/51682-getting-info-about-drives-vbnet/
    Private Function setPrefix(ByVal size As Long) As String
        Dim TotalString As String = ""
        For Prefix As Integer = 0 To 4

            If size < 1024 Or Prefix = 4 Then
                TotalString = size
                Select Case Prefix
                    Case 0
                        TotalString &= " B"
                    Case 1
                        TotalString &= " KB"
                    Case 2
                        TotalString &= " MB"
                    Case 3
                        TotalString &= " GB"
                    Case 4
                        TotalString &= " TB"
                End Select
                Exit For
            Else
                size /= 1024
            End If
        Next
        Return TotalString
    End Function
    '------------------------------------------------------------------------------------------
    '  Button Functions
    '------------------------------------------------------------------------------------------
    'Display app exit confirmation
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        If MsgBox("Are you sure you want to exit  ComputerInfo?", MsgBoxStyle.OkCancel, " ComputerInfo") = MsgBoxResult.Ok Then
            StopWord() 'Terminate Word process
            DeleteHeader() 'Delete ICHeader.png in current directory
            Application.Exit()
        End If
    End Sub

    'Send Email
    Private Sub EmailButton_Click(sender As Object, e As EventArgs) Handles EmailButton.Click
        Dim attachment As String
        attachment = Application.ExecutablePath & GetComputerName() + ".pdf"

        If System.IO.File.Exists(attachment) = False Then
            MessageBox.Show("Please export into PDF first and then try again.", " ComputerInfo")

        Else
            SendEmail()
        End If
    End Sub

    'Open exported PDF file
    Private Sub OpenExistingFileButton_Click(sender As Object, e As EventArgs) Handles OpenPDFButton.Click
        Dim FILE_NAME As String = Application.ExecutablePath & GetComputerName() + ".pdf"

        If System.IO.File.Exists(FILE_NAME) = True Then
            Process.Start(FILE_NAME)
        Else
            MessageBox.Show("Please export into PDF first and then try again.", " ComputerInfo")
        End If
    End Sub

    'Save all computer information onto PDF file
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles SaveAsPDFButton.Click
        ExportAppList() 'Exporting list of installed app at first place
        StopNotepad() 'Terminate Notepad after complete export process
        Dim fileReader As String
        Dim WordType As Type
        WordType = Type.GetTypeFromProgID("Word.Application") 'Get the Word properties
        fileReader = My.Computer.FileSystem.ReadAllText(Application.ExecutablePath & GetComputerName() + ".txt") 'Prepare reading text location

        If WordType Is Nothing Then 'Check if Word is installed
            MessageBox.Show("Please install Microsoft Word and then try again.", " ComputerInfo")

        Else
            'Check if the text field is empty
            If (EngineerNameInput.Text.Trim = String.Empty) Or (EmployeeNameInput.Text.Trim = String.Empty) Or (EmployeeIDInput.Text.Trim = String.Empty) Or (ContactNumberInput.Text.Trim = String.Empty) Or (LANIDInput.Text.Trim = String.Empty) Or (DivisionInput.Text.Trim = String.Empty) Or (DepartmentInput.Text.Trim = String.Empty) Or (LocationInput.Text.Trim = String.Empty) Or (CostCentreInput01.Text.Trim = String.Empty) Or
           (String.IsNullOrEmpty(EngineerNameInput.Text.Trim)) Or (String.IsNullOrEmpty(EmployeeNameInput.Text.Trim)) Or (String.IsNullOrEmpty(EmployeeIDInput.Text.Trim)) Or (String.IsNullOrEmpty(ContactNumberInput.Text.Trim)) Or (String.IsNullOrEmpty(LANIDInput.Text.Trim)) Or (String.IsNullOrEmpty(DivisionInput.Text.Trim)) Or (String.IsNullOrEmpty(DepartmentInput.Text.Trim)) Or (String.IsNullOrEmpty(LocationInput.Text.Trim)) Or (String.IsNullOrEmpty(CostCentreInput01.Text.Trim)) Then
                MessageBox.Show("Please fill in all the required details.", " ComputerInfo")
            Else
                Try
                    Dim oWord As Word.Application
                    Dim oDoc As Word.Document
                    Dim oTablePCInfo As Word.Table
                    Dim oTableEmployeeInfo As Word.Table
                    Dim oPara1 As Word.Paragraph, oPara2 As Word.Paragraph
                    Dim oPara3 As Word.Paragraph, oPara4 As Word.Paragraph
                    Dim oPara5 As Word.Paragraph, oPara6 As Word.Paragraph
                    Dim oPara7 As Word.Paragraph, oPara8 As Word.Paragraph
                    Dim oPara9 As Word.Paragraph, oPara10 As Word.Paragraph
                    ExportHeader()

                    'Start Word and open the document template.
                    oWord = CreateObject("Word.Application")
                    oWord.Visible = False
                    oDoc = oWord.Documents.Add

                    'Set the page margin
                    oDoc.PageSetup.PageWidth = 500
                    oDoc.PageSetup.PageHeight = 750
                    oDoc.PageSetup.RightMargin = 25.5
                    oDoc.PageSetup.LeftMargin = 25.5
                    oDoc.PageSetup.TopMargin = 25.5
                    oDoc.PageSetup.BottomMargin = 28.4

                    'Add paragraph
                    Dim para As Word.Paragraph = oDoc.Paragraphs.Add()
                    Dim PictureLocation As String

                    PictureLocation = Application.StartupPath & "\ICHeader.png"
                    para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    para.Range.InlineShapes.AddPicture(PictureLocation)

                    'Add paragraph
                    oPara1 = oDoc.Content.Paragraphs.Add
                    oPara1.Range.Text = "Date: " + GetDate() + ""
                    oPara1.Range.Font.Bold = False
                    oPara1.Format.SpaceAfter = 6    '6 pt spacing after paragraph.
                    oPara1.Range.InsertParagraphAfter()

                    'Create table
                    oTablePCInfo = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 10, 4)
                    oTablePCInfo.Borders.Enable = True 'All borders enabled on the table
                    oTablePCInfo.Cell(1, 1).Range.Text = "Computer Name"
                    oTablePCInfo.Cell(2, 1).Range.Text = "User Name"
                    oTablePCInfo.Cell(3, 1).Range.Text = "Operating System"
                    oTablePCInfo.Cell(4, 1).Range.Text = "Service Pack"
                    oTablePCInfo.Cell(5, 1).Range.Text = "Brand Name"
                    oTablePCInfo.Cell(6, 1).Range.Text = "Model"
                    oTablePCInfo.Cell(7, 1).Range.Text = "Processors"
                    oTablePCInfo.Cell(8, 1).Range.Text = "Memory (RAM)"
                    oTablePCInfo.Cell(9, 1).Range.Text = "Serial Number"
                    oTablePCInfo.Cell(10, 1).Range.Text = "Default Printer"

                    oTablePCInfo.Cell(1, 3).Range.Text = "IP Address"
                    oTablePCInfo.Cell(2, 3).Range.Text = "MAC Address"
                    oTablePCInfo.Cell(3, 3).Range.Text = "Primary DNS"
                    oTablePCInfo.Cell(4, 3).Range.Text = "Secondary DNS"
                    oTablePCInfo.Cell(5, 3).Range.Text = "Drive C"
                    oTablePCInfo.Cell(6, 3).Range.Text = "Drive D"
                    oTablePCInfo.Cell(7, 3).Range.Text = "Image Version"
                    oTablePCInfo.Cell(8, 3).Range.Text = "Drive C Encryption"
                    oTablePCInfo.Cell(9, 3).Range.Text = "Drive D Encryption"
                    oTablePCInfo.Cell(10, 3).Range.Text = "USB Status"

                    'Set the text range of the table cell into bold
                    oTablePCInfo.Cell(1, 1).Range.Font.Bold = True
                    oTablePCInfo.Cell(2, 1).Range.Font.Bold = True
                    oTablePCInfo.Cell(3, 1).Range.Font.Bold = True
                    oTablePCInfo.Cell(4, 1).Range.Font.Bold = True
                    oTablePCInfo.Cell(5, 1).Range.Font.Bold = True
                    oTablePCInfo.Cell(6, 1).Range.Font.Bold = True
                    oTablePCInfo.Cell(7, 1).Range.Font.Bold = True
                    oTablePCInfo.Cell(8, 1).Range.Font.Bold = True
                    oTablePCInfo.Cell(9, 1).Range.Font.Bold = True
                    oTablePCInfo.Cell(10, 1).Range.Font.Bold = True
                    oTablePCInfo.Cell(1, 3).Range.Font.Bold = True
                    oTablePCInfo.Cell(2, 3).Range.Font.Bold = True
                    oTablePCInfo.Cell(3, 3).Range.Font.Bold = True
                    oTablePCInfo.Cell(4, 3).Range.Font.Bold = True
                    oTablePCInfo.Cell(5, 3).Range.Font.Bold = True
                    oTablePCInfo.Cell(6, 3).Range.Font.Bold = True
                    oTablePCInfo.Cell(7, 3).Range.Font.Bold = True
                    oTablePCInfo.Cell(8, 3).Range.Font.Bold = True
                    oTablePCInfo.Cell(9, 3).Range.Font.Bold = True
                    oTablePCInfo.Cell(10, 3).Range.Font.Bold = True

                    'Set the cell range of the table cell into light gray
                    oTablePCInfo.Cell(1, 1).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTablePCInfo.Cell(2, 1).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTablePCInfo.Cell(3, 1).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTablePCInfo.Cell(4, 1).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTablePCInfo.Cell(5, 1).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTablePCInfo.Cell(6, 1).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTablePCInfo.Cell(7, 1).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTablePCInfo.Cell(8, 1).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTablePCInfo.Cell(9, 1).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTablePCInfo.Cell(10, 1).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTablePCInfo.Cell(1, 3).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTablePCInfo.Cell(2, 3).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTablePCInfo.Cell(3, 3).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTablePCInfo.Cell(4, 3).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTablePCInfo.Cell(5, 3).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTablePCInfo.Cell(6, 3).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTablePCInfo.Cell(7, 3).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTablePCInfo.Cell(8, 3).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTablePCInfo.Cell(9, 3).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTablePCInfo.Cell(10, 3).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)

                    'Write results
                    oTablePCInfo.Cell(1, 2).Range.Text = ComputerNameResult.Text
                    oTablePCInfo.Cell(2, 2).Range.Text = UserNameResult.Text
                    oTablePCInfo.Cell(3, 2).Range.Text = OSResult.Text
                    oTablePCInfo.Cell(4, 2).Range.Text = SPResult.Text
                    oTablePCInfo.Cell(5, 2).Range.Text = BrandNameResult.Text
                    oTablePCInfo.Cell(6, 2).Range.Text = ModelResult.Text
                    oTablePCInfo.Cell(7, 2).Range.Text = ProcessorsResult.Text
                    oTablePCInfo.Cell(8, 2).Range.Text = RAMResult.Text
                    oTablePCInfo.Cell(9, 2).Range.Text = SerialNumberResult.Text
                    oTablePCInfo.Cell(10, 2).Range.Text = DefaultPrinterResult.Text

                    oTablePCInfo.Cell(1, 4).Range.Text = IPAddressResult.Text
                    oTablePCInfo.Cell(2, 4).Range.Text = MACResult.Text
                    oTablePCInfo.Cell(3, 4).Range.Text = PrimaryDNSResult.Text
                    oTablePCInfo.Cell(4, 4).Range.Text = SecondaryDNSResult.Text
                    oTablePCInfo.Cell(5, 4).Range.Text = ExportDriveC()
                    oTablePCInfo.Cell(6, 4).Range.Text = ExportDriveD()
                    oTablePCInfo.Cell(7, 4).Range.Text = GetImageVersion()
                    oTablePCInfo.Cell(8, 4).Range.Text = GetDriveCEncryptionStatus()
                    oTablePCInfo.Cell(9, 4).Range.Text = GetDriveDEncryptionStatus()
                    oTablePCInfo.Cell(10, 4).Range.Text = GetUSBState()

                    'Add paragraph
                    oPara2 = oDoc.Content.Paragraphs.Add
                    oPara2.Range.Text = ""
                    oPara2.Range.Font.Bold = False
                    oPara2.Format.SpaceAfter = 20    '20 pt spacing after paragraph.
                    oPara2.Range.InsertParagraphAfter()

                    'Add paragraph
                    oPara3 = oDoc.Content.Paragraphs.Add
                    oPara3.Range.Text = "Inventory Information:"
                    oPara3.Range.Font.Bold = False
                    oPara3.Format.SpaceAfter = 9    '20 pt spacing after paragraph.
                    oPara3.Range.InsertParagraphAfter()

                    'Create table
                    oTableEmployeeInfo = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 4)
                    oTableEmployeeInfo.Borders.Enable = True
                    oTableEmployeeInfo.Cell(1, 1).Range.Text = "Employee Name"
                    oTableEmployeeInfo.Cell(2, 1).Range.Text = "Employee ID"
                    oTableEmployeeInfo.Cell(3, 1).Range.Text = "Contact Number"
                    oTableEmployeeInfo.Cell(4, 1).Range.Text = "LAN ID"
                    oTableEmployeeInfo.Cell(5, 1).Range.Text = "Division"
                    oTableEmployeeInfo.Cell(1, 3).Range.Text = "Department"
                    oTableEmployeeInfo.Cell(2, 3).Range.Text = "Location"
                    oTableEmployeeInfo.Cell(3, 3).Range.Text = "Cost Centre"
                    oTableEmployeeInfo.Cell(4, 3).Range.Text = "Engineer Name"
                    oTableEmployeeInfo.Cell(5, 3).Range.Text = "Checklist Type"

                    'Set the text range of the table cell into bold
                    oTableEmployeeInfo.Cell(1, 1).Range.Font.Bold = True
                    oTableEmployeeInfo.Cell(2, 1).Range.Font.Bold = True
                    oTableEmployeeInfo.Cell(3, 1).Range.Font.Bold = True
                    oTableEmployeeInfo.Cell(4, 1).Range.Font.Bold = True
                    oTableEmployeeInfo.Cell(5, 1).Range.Font.Bold = True
                    oTableEmployeeInfo.Cell(1, 3).Range.Font.Bold = True
                    oTableEmployeeInfo.Cell(2, 3).Range.Font.Bold = True
                    oTableEmployeeInfo.Cell(3, 3).Range.Font.Bold = True
                    oTableEmployeeInfo.Cell(4, 3).Range.Font.Bold = True
                    oTableEmployeeInfo.Cell(5, 3).Range.Font.Bold = True

                    'Set the cell range of the table cell into light gray
                    oTableEmployeeInfo.Cell(1, 1).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTableEmployeeInfo.Cell(2, 1).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTableEmployeeInfo.Cell(3, 1).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTableEmployeeInfo.Cell(4, 1).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTableEmployeeInfo.Cell(5, 1).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTableEmployeeInfo.Cell(1, 3).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTableEmployeeInfo.Cell(2, 3).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTableEmployeeInfo.Cell(3, 3).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTableEmployeeInfo.Cell(4, 3).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)
                    oTableEmployeeInfo.Cell(5, 3).Range.Shading.BackgroundPatternColor = RGB(211, 211, 211)

                    'Write results
                    oTableEmployeeInfo.Cell(1, 2).Range.Text = GetEmployeeName()
                    oTableEmployeeInfo.Cell(2, 2).Range.Text = GetEmployeeID()
                    oTableEmployeeInfo.Cell(3, 2).Range.Text = GetContactNumber()
                    oTableEmployeeInfo.Cell(4, 2).Range.Text = GetLANID()
                    oTableEmployeeInfo.Cell(5, 2).Range.Text = GetDivision()
                    oTableEmployeeInfo.Cell(1, 4).Range.Text = GetDepartment()
                    oTableEmployeeInfo.Cell(2, 4).Range.Text = GetLocation()
                    oTableEmployeeInfo.Cell(3, 4).Range.Text = GetCallCentre()
                    oTableEmployeeInfo.Cell(4, 4).Range.Text = GetEngineerName()
                    oTableEmployeeInfo.Cell(5, 4).Range.Text = GetChecklistType()

                    'Add paragraph
                    oPara4 = oDoc.Content.Paragraphs.Add
                    oPara4.Range.Text = ""
                    oPara4.Range.Font.Bold = False
                    oPara4.Format.SpaceAfter = 120    '20 pt spacing after paragraph.
                    oPara4.Range.InsertParagraphAfter()

                    oPara5 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
                    oPara5.Range.Text = "The following app were installed on this computer:"
                    oPara5.Format.SpaceBefore = 1
                    oPara5.Format.SpaceAfter = 1
                    oPara5.Range.InsertParagraphAfter()

                    oPara6 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
                    oPara6.Format.SpaceBefore = 1
                    oPara6.Range.Text = " "
                    oPara6.Format.SpaceAfter = 1
                    oPara6.Range.InsertParagraphAfter()

                    'Writing installed app into Word
                    oPara7 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
                    oPara7.Format.SpaceBefore = 1
                    oPara7.Range.Text = fileReader
                    oPara7.Format.SpaceAfter = 1
                    oPara7.Range.InsertParagraphAfter()

                    oPara8 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
                    oPara8.Format.SpaceBefore = 1
                    Form2.ShowAppList()
                    oPara8.Range.Text = "Total installed App: " + Form2.AppList.Items.Count.ToString()
                    oPara8.Format.SpaceAfter = 1
                    oPara8.Range.InsertParagraphAfter()

                    oPara9 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
                    oPara9.Format.SpaceBefore = 1
                    oPara9.Range.Text = " "
                    oPara9.Format.SpaceAfter = 1
                    oPara9.Range.InsertParagraphAfter()

                    'Add paragraph
                    oPara10 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
                    oPara10.Range.Text = "This inventory checklist is automatically generated by  ComputerInfo."
                    oPara10.Format.SpaceBefore = 20
                    oPara10.Format.SpaceAfter = 6
                    oPara10.Range.InsertParagraphAfter()

                    'Programatically write onto Word Documents and then convert into PDF
                    Dim newdoc As Word.Document
                    Dim SavePath As String
                    SavePath = Application.ExecutablePath & GetComputerName() + ".pdf"
                    newdoc = oDoc
                    newdoc.SaveAs2(SavePath, Word.WdSaveFormat.wdFormatPDF)
                    StopWord() 'Terminate Microsoft Word after saving completes

                    'Display message
                    MessageBox.Show("PDF export is completed.", " ComputerInfo")

                Catch ex As Exception
                    MessageBox.Show("Error exporting to PDF format. Technical information are as follows:" & vbCrLf & ex.Message, " ComputerInfo")
                End Try
            End If
        End If
    End Sub

    'Save all computer information onto CSV file
    Private Sub SaveAsCSVButton_Click(sender As Object, e As EventArgs) Handles SaveAsCSVButton.Click
        If (EngineerNameInput.Text.Trim = String.Empty) Or (EmployeeNameInput.Text.Trim = String.Empty) Or (EmployeeIDInput.Text.Trim = String.Empty) Or (ContactNumberInput.Text.Trim = String.Empty) Or (LANIDInput.Text.Trim = String.Empty) Or (DivisionInput.Text.Trim = String.Empty) Or (DepartmentInput.Text.Trim = String.Empty) Or (LocationInput.Text.Trim = String.Empty) Or (CostCentreInput01.Text.Trim = String.Empty) Or
           (String.IsNullOrEmpty(EngineerNameInput.Text.Trim)) Or (String.IsNullOrEmpty(EmployeeNameInput.Text.Trim)) Or (String.IsNullOrEmpty(EmployeeIDInput.Text.Trim)) Or (String.IsNullOrEmpty(ContactNumberInput.Text.Trim)) Or (String.IsNullOrEmpty(LANIDInput.Text.Trim)) Or (String.IsNullOrEmpty(DivisionInput.Text.Trim)) Or (String.IsNullOrEmpty(DepartmentInput.Text.Trim)) Or (String.IsNullOrEmpty(LocationInput.Text.Trim)) Or (String.IsNullOrEmpty(CostCentreInput01.Text.Trim)) Then
            MessageBox.Show("Please fill in all the required details.", " ComputerInfo")
        Else
            Try
                Dim csvFile As String = Application.ExecutablePath & GetComputerName() + ".csv"
                Dim outFile As IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(csvFile, False)

                outFile.WriteLine(" ,Desktop Code,Status,Location,Division,Department,New Dept CC,DO Number,Floor,User Name,Manufacturer,Model Description,Machine Type,Serial Number,OS Type,Domain,IP Address,Workstation Name,Workstation Label,Monitor Manufacturer,Monitor Model Description,Monitor Machine Type,Monitor Serial Number,Monitor Tech Type,Monitor Tech Size,MyKad Reader Type,MyKad Reader Serial Number,USB Blocked,SMILE Case No,Engineer Name,Engineer Contact No,Remarks")
                outFile.WriteLine(" ," + " ," + " ," + GetLocation() + "," + GetDivision() + "," + GetDepartment() + "," + GetCallCentre() + "," + " ," + " ," + GetEmployeeName() + "," + GetBrandName() + "," + GetModel() + "," + " ," + GetSerialNumber() + "," + GetOperatingSystem() + "," + " ," + GetIPAddressIPv4() + "," + GetComputerName() + "," + " ," + " ," + " ," + " ," + " ," + " ," + " ," + " ," + " ," + " ," + " ," + GetEngineerName() + "," + " ," + " ,")
                outFile.Close()
                MessageBox.Show("CSV export is completed.", " ComputerInfo")

            Catch ex As Exception
                MessageBox.Show("Error exporting to CSV format. Technical information are as follows:" & vbCrLf & ex.Message, " ComputerInfo")
            End Try
        End If
    End Sub

    'Enforce Contact Number column's numberic-only
    Private Sub ContactNumberInput_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ContactNumberInput.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            ContactNumberInput.Text = " "
            e.Handled = True
        End If
    End Sub

    'Update results on button click
    Private Sub RefreshButton_Click(sender As Object, e As EventArgs) Handles RefreshButton.Click
        'Right-hand Part
        ShowIPAddress() 'Display IP Address
        ShowMACAddress() 'Get MAC Address
        ShowPrimaryDNSAddress() 'Display primary DNS Address
        ShowSecondaryDNSAddress() 'Display secondary DNS Address
        DisplayHardDiskSpace() 'Display Hard Disk Space
        ShowDriveCEncryptionStatus() 'Display Drive C encryption status
        ShowDriveDEncryptionStatus() 'Display Drive D encryption status
        ShowUSBState() 'Display the status of USB

        MessageBox.Show("The refresh is completed.", " ComputerInfo")
    End Sub

    'Call out View Installed App form
    Private Sub ViewInstalledAppButton_Click(sender As Object, e As EventArgs) Handles ViewInstalledAppButton.Click
        Form2.ShowDialog()
    End Sub

    'Call out About App form
    Private Sub AboutAppButton_Click(sender As Object, e As EventArgs) Handles AboutAppButton.Click
        Form3.ShowDialog()
    End Sub

    'Enforce Cost Centre input on every column maximum length is 3
    Private Sub CostCentreInput01_TextChanged(sender As Object, e As EventArgs) Handles CostCentreInput01.TextChanged, CostCentreInput02.TextChanged
        Dim txtbox As Control = DirectCast(sender, TextBox)

        If txtbox.Text.Length = 3 Then
            GetNextControl(ActiveControl, True).Focus()

        End If
    End Sub

    'Autofocus for Cost Centre Input
    Private Sub CostCentreInput01_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CostCentreInput01.KeyPress
        If Char.IsNumber(e.KeyChar) = False Then
            e.Handled = True
        End If
    End Sub

    'Autofocus for Cost Centre Input
    Private Sub CostCentreInput02_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CostCentreInput02.KeyPress
        If Char.IsNumber(e.KeyChar) = False Then
            e.Handled = True
        End If
    End Sub

    'Autofocus for Cost Centre Input
    Private Sub CostCentreInput03_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CostCentreInput03.KeyPress
        If Char.IsNumber(e.KeyChar) = False Then
            e.Handled = True
        End If
    End Sub
End Class
