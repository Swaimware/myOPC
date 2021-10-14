Imports System.Windows.Forms
Imports OPCAutomation

''' <summary>
''' Name:        myOPC
''' Company:     SwaimWare
''' Author:      Kurt Swaim
''' Description: 
''' 
''' myOPC.dll allows the user of this object to create an OPC connection 
''' to either KepServerEX V5, Cognex, or RSLinx OPC server.  Initial properties 
''' must be setup before starting the OPC connection.  Specifically the 
''' properties NumItems, TopicName (RSLinx), ChannelName (Kepware), 
''' DeviceName (Kepware), CameraName (Cognex) GroupName, DeadBand, UpdateRate, 
''' OPCServerType, ServerType, SilentMode.  Along with 
''' these properties one public field item array OPCItemNames() has to 
''' be filled with the tags in the PLC to be read/wrote to.  Note that 
''' the size of the array is dictated by the property NumItems.  In addition 
''' two properties are defaulted.  The property SilentMode if set to true will 
''' inhibit message boxes with fault descriptions, the default is false indicating 
''' the message boxes will appear during errors.  When the setup values 
''' are set use the StartOPC function to get the OPCItemValues() to start holding 
''' data pointed at.  If a TRUE is returned with StartOPC means no errors occurred.  
''' If a FALSE is returned then either the initial data had noting in it or was 
''' invalid or the physical connection is bad.  The public field item OPCItemActive() 
''' will indicate if the OPCItemName() was found and will be true if all went 
''' well and the tag is active.  The public field item is OPCItemQuality() will be 
''' updated at every read interval with "Good" "Uncertain" or "Bad" indicating the link 
''' quality.  The OPCwrite function allows for individual OPCItemNames() tags on the 
''' PLC to written to with a new value.  If the function returns true the value was 
''' written to successfully.  Lastly be a good programmer and call the Close method 
''' when done with the application.  Enjoy.
''' </summary>
''' <remarks></remarks>
Public Class myOPC

#Region "Field Items"

    ' Private Field Items to setup the server
    Private _DoubleShot As Boolean
    Private _NumItems As Integer
    Private _ItemCount As Integer
    Private _TopicName As String
    Private _ChannelName As String
    Private _DeviceName As String
    Private _CameraName As String
    Private _GroupName As String
    Private _DeadBand As Integer
    Private _UpdateRate As Integer
    Private _ActiveState As Boolean
    Private _SilentMode As Boolean = False
    Private _AsynchMode As Boolean = True
    ' Private Field Items Describing OPC name
    Private Const _RSLinxOPCserver As String = "RSLinx OPC Server"
    Private Const _KepwareOPCserver As String = "Kepware.KEPServerEX.V5"
    Private Const _CognexOPCserver As String = "Cognex In-Sight OPC Server"
    Private _OPCserver As String = _RSLinxOPCserver
    Private _OPCServerType As Integer = 1
    ' Private Field Items used with OPC communication methods
    Private _ItemServerHandles As System.Array
    Private _AddItemServerErrors As System.Array
    Private _RemoveItemServerErrors As System.Array
    Private _SyncItemValues As System.Array
    Private _SyncItemServerErrors As System.Array
    Private _SyncItemQuality As System.Array
    Private _SyncItemTimeStamp As System.Array
    Private _ActiveItemErrors As System.Array
    Private _SyncItemServerHandles() As Integer
    Private _ActiveItemServerHandles() As Int32
    Private _RemoveItemServerHandles() As Int32
    Private _OPCItemIDs() As String
    Private _ClientHandles() As Integer
    ' Public Field Items for the user of objects use
    ''' <summary>
    ''' Array of the names of the PLC tags that the OPC server will 
    ''' read/write data.  The array size is dictated by the property NumItems.
    ''' </summary>
    ''' <remarks></remarks>
    Public OPCItemNames() As Object
    ''' <summary>
    ''' Array of Bools indicating if the corresponding OPCItemNames() tags are 
    ''' active.  True being active. The array size is dictated by the property NumItems.
    ''' </summary>
    ''' <remarks></remarks>
    Public OPCItemActive() As Boolean
    ''' <summary>
    ''' Array of values from the PLC tags pointed to from OPCItemNames coming in 
    ''' from the OPC server.  The array size is dictated by the property NumItems.
    ''' </summary>
    ''' <remarks></remarks>
    Public OPCItemValues() As Object
    ''' <summary>
    ''' Array of quality indicating strings specifying how the communication is between 
    ''' the OPC server and the PLC tags. The array size is dictated by the property NumItems.
    ''' </summary>
    ''' <remarks></remarks>
    Public OPCItemQuality() As String
    ''' <summary>
    ''' Array of time stamps indicating when the data was collected.  The array size is dictated 
    ''' by the property NumItems.
    ''' </summary>
    ''' <remarks></remarks>
    Public OPCItemTimeStamp() As Date
    ' Define the OPC Automation objects
    Private WithEvents _ConnectedOPCServer As OPCServer
    Private WithEvents _ConnectedGroup As OPCGroup

#End Region

#Region "Properties"

    ''' <summary>
    ''' Number of OPC items to be created.  
    ''' Because changing this value effects the array size 
    ''' they are re-dimensioned.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property NumItems() As Integer
        Get
            NumItems = _NumItems
        End Get
        Set(ByVal value As Integer)
            _NumItems = value
            ReDim _SyncItemServerHandles(_NumItems)
            ReDim _ActiveItemServerHandles(1)
            ReDim _RemoveItemServerHandles(_NumItems)
            ReDim _OPCItemIDs(_NumItems)
            ReDim _ClientHandles(_NumItems)
            ReDim OPCItemNames(_NumItems)
            ReDim OPCItemActive(_NumItems)
            ReDim OPCItemValues(_NumItems)
            ReDim OPCItemQuality(_NumItems)
            ReDim OPCItemTimeStamp(_NumItems)
            _ItemCount = _NumItems
        End Set
    End Property

    ''' <summary>
    ''' The topic name as created in RSLinx.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TopicName() As String
        Get
            TopicName = _TopicName
        End Get
        Set(ByVal value As String)
            _TopicName = value
        End Set
    End Property

    ''' <summary>
    ''' The channel name as created in Kepware.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ChannelName() As String
        Get
            ChannelName = _ChannelName
        End Get
        Set(ByVal value As String)
            _ChannelName = value
        End Set
    End Property

    ''' <summary>
    ''' The device name as created in Kepware.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DeviceName() As String
        Get
            DeviceName = _DeviceName
        End Get
        Set(ByVal value As String)
            _DeviceName = value
        End Set
    End Property

    ''' <summary>
    ''' The Cognex camera name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CameraName() As String
        Get
            CameraName = _CameraName
        End Get
        Set(ByVal value As String)
            _CameraName = value
        End Set
    End Property

    ''' <summary>
    ''' The OPC group name that will house the OPC items.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property GroupName() As String
        Get
            GroupName = _GroupName
        End Get
        Set(ByVal value As String)
            _GroupName = value
        End Set
    End Property

    ''' <summary>
    ''' The OPC Dead Band determines how much in percent the item 
    ''' value has to change before it gets updated.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DeadBand() As Integer
        Get
            DeadBand = _DeadBand
        End Get
        Set(ByVal value As Integer)
            If value >= 100 Then
                _DeadBand = 99
            ElseIf (value < 0) Then
                _DeadBand = 0
            Else
                _DeadBand = value
            End If
        End Set
    End Property

    ''' <summary>
    ''' The OPC data item value read rate.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property UpdateRate() As Integer
        Get
            UpdateRate = _UpdateRate
        End Get
        Set(ByVal value As Integer)
            _UpdateRate = value
        End Set
    End Property

    ''' <summary>
    ''' In silent mode the message boxes are inhibited.  
    ''' Set to TRUE to silence message boxes
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SilentMode() As Boolean
        Get
            SilentMode = _SilentMode
        End Get
        Set(ByVal value As Boolean)
            _SilentMode = value
        End Set
    End Property

    ''' <summary>
    ''' Determines if Kepware or RSLinx is the OPC server to be used.  
    ''' 1 = RSLinx 2 = Kepware  3 = Cognex
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property OPCServerType() As Integer
        Get
            OPCServerType = _OPCServerType
        End Get
        Set(ByVal value As Integer)
            If value < 1 Or value > 3 Then
                _OPCServerType = 1
            Else
                _OPCServerType = value
            End If
            Select Case _OPCServerType
                Case 1
                    _OPCserver = _RSLinxOPCserver
                Case 2
                    _OPCserver = _KepwareOPCserver
                Case 3
                    _OPCserver = _CognexOPCserver
            End Select
        End Set
    End Property

    ''' <summary>
    ''' If set to true the group will be placed in asynchronous updates via the DataChange event. 
    ''' Otherwise a synchronous read will be required to read data.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property AsynchMode() As Boolean
        Get
            AsynchMode = _AsynchMode
        End Get
        Set(ByVal value As Boolean)
            _AsynchMode = value
        End Set
    End Property

    ''' <summary>
    ''' Returns the OPCServerState
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ServerState As OPCServerState
        Get
            ServerState = _ConnectedOPCServer.ServerState
        End Get
    End Property

#End Region

#Region "The Methods"

    ''' <summary>
    ''' Default constructor.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

    End Sub

    ''' <summary>
    ''' Creates the OPC connection, Groups, Items then activates and reads item values 
    ''' as long as no errors occur along the way.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function StartOPC() As Boolean
        Try
            ' First verify that all of the parameters have been set up
            ' before starting OPC connection.
            If Not _DeadBand = Nothing _
            Or Not _UpdateRate = Nothing _
            Or Not _GroupName = Nothing _
            Or Not _NumItems = Nothing _
            Or ((Not _TopicName = Nothing) Or (Not _CameraName = Nothing) Or ((Not _ChannelName = Nothing) And (Not _DeviceName = Nothing))) Then
                For j As Integer = 1 To _NumItems
                    If OPCItemNames(j) Is Nothing Then
                        If _SilentMode = False Then
                            MessageBox.Show("Set the OPC parameters before opening connection", "Error", MessageBoxButtons.OK)
                        End If
                        Return False
                    End If
                Next
                ' Create a new OPC Server object
                _ConnectedOPCServer = New OPCServer
                ' Attempt to connect with the server
                _ConnectedOPCServer.Connect(_OPCserver, "")
                ' Set the desire active state for the group
                _ConnectedOPCServer.OPCGroups.DefaultGroupIsActive = True
                ' Set the desired percent dead band
                _ConnectedOPCServer.OPCGroups.DefaultGroupDeadband = _DeadBand
                ' Add the group
                _ConnectedGroup = _ConnectedOPCServer.OPCGroups.Add(_GroupName)
                ' Set the update rate for the group
                _ConnectedGroup.UpdateRate = _UpdateRate
                ' Mark this group to receive asynchronous updates via the DataChange event.
                _ConnectedGroup.IsSubscribed = _AsynchMode
                ' Setting the '.DefaultIsActive' property forces all items we are about to
                ' add to the group to be added in a non active state.
                _ConnectedGroup.OPCItems.DefaultIsActive = False
                ' Assemble the OPC Item IDs
                For i As Integer = 1 To _NumItems
                    ' Define item name based on which OPC server utilized.
                    Select Case _OPCServerType
                        Case 1 'RSLinx OPC Server
                            _OPCItemIDs(i) = String.Format("[{0}]{1}", _TopicName, CStr(OPCItemNames(i)))
                        Case 2 'Kepware OPC Server
                            _OPCItemIDs(i) = String.Format("{0}.{1}.{2}", _ChannelName, _DeviceName, CStr(OPCItemNames(i)))
                        Case 3 'Cognex OPC Server
                            _OPCItemIDs(i) = String.Format("{0}.{1}", _CameraName, CStr(OPCItemNames(i)))
                    End Select
                    ' Define the Client Handle
                    _ClientHandles(i) = i
                Next
                ' Add the item
                _ConnectedGroup.OPCItems.AddItems(_ItemCount, _OPCItemIDs, _ClientHandles, _ItemServerHandles, _AddItemServerErrors)
                ' For the items added without errors make them active 
                For i = 1 To _NumItems
                    If _AddItemServerErrors(i) <> 0 Then
                        ' Set the active type desired.
                        _ActiveState = False
                        ' Get the Servers handle for the desired item.  The server handles
                        ' were returned in add item subroutine.
                        _ActiveItemServerHandles(1) = _ItemServerHandles(i)
                        ' Invoke the SetActive operation on the OPC item collection interface
                        _ConnectedGroup.OPCItems.SetActive(1, _ActiveItemServerHandles, _ActiveState, _ActiveItemErrors)
                        ' Inform user of class is not active because of faults.
                        OPCItemActive(i) = False
                    Else
                        ' Set the active type desired.
                        _ActiveState = True
                        ' Get the Servers handle for the desired item.  The server handles
                        ' were returned in add item subroutine.
                        _ActiveItemServerHandles(1) = _ItemServerHandles(i)
                        ' Invoke the SetActive operation on the OPC item collection interface
                        _ConnectedGroup.OPCItems.SetActive(1, _ActiveItemServerHandles, _ActiveState, _ActiveItemErrors)
                        ' If an error occurred during activation then deactivate the user info.
                        OPCItemActive(i) = True
                    End If
                    ' Get the Servers handle for the desired item.  The server handles were
                    ' returned in add item subroutine.
                    _SyncItemServerHandles(i) = _ItemServerHandles(i)
                    _RemoveItemServerHandles(i) = _ItemServerHandles(i)
                Next
                ' Invoke the SyncRead operation.  Remember this call will wait until
                ' completion. The source flag in this case, 'OPCDevice' , is set to
                ' read from device which may take some time.
                _ConnectedGroup.SyncRead(OPCDataSource.OPCDevice, _ItemCount, _SyncItemServerHandles, _SyncItemValues, _SyncItemServerErrors)
                ' Collect data read and if any errors occurred or the item is not
                ' active then zero out the value sent to the user.
                For i = 1 To NumItems
                    If OPCItemActive(i) Then
                        If _SyncItemServerErrors(i) = 0 Then
                            OPCItemValues(i) = _SyncItemValues(i)
                        Else
                            OPCItemValues(i) = 0
                            Return False
                        End If
                    Else
                        OPCItemValues(i) = 0
                    End If
                Next
                Return True
            Else
                If _SilentMode = False Then
                    MessageBox.Show("Set the OPC parameters before opening connection", "Error", MessageBoxButtons.OK)
                End If
                Return False
            End If
        Catch ex As Exception
            ' Error handling
            _ConnectedOPCServer = Nothing
            If _SilentMode = False Then
                MessageBox.Show("OPC server connect failed with exception: " + ex.Message, "myOPC Exception", MessageBoxButtons.OK)
            End If
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Check the Qualities for each item returned here.  The actual contents of the
    ''' quality field can contain bit field data which can provide specific
    ''' error conditions.  Normally if everything is OK then the quality will
    ''' contain the 0xC0
    ''' </summary>
    ''' <param name="Quality"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Shared Function OPCquality(ByVal Quality As OPCAutomation.OPCQuality) As String
        If Quality = OPCAutomation.OPCQuality.OPCQualityGood Then
            Return "Good"
        ElseIf Quality = OPCAutomation.OPCQuality.OPCQualityUncertain Then
            Return "Uncertain"
        Else
            Return "Bad"
        End If
    End Function

    Protected Overridable Sub _ConnectedOPCServer_ServerShutdown(ByVal MSG As String) Handles _ConnectedOPCServer.ServerShutDown
        ' Add your crap here

    End Sub

    Private Sub _ConnectedGroup_DataChange(ByVal TransactionID As Integer, ByVal NumItems As Integer, ByRef ClientHandles As System.Array, ByRef ItemValues As System.Array, ByRef Qualities As System.Array, ByRef TimeStamps As System.Array) Handles _ConnectedGroup.DataChange
        Try
            Dim i As Integer
            For i = 1 To NumItems
                ' Use the 'ClientHandles' array returned by the server to pull out the
                ' index number of the control to update and load the value.
                OPCItemValues(ClientHandles(i)) = ItemValues(i)
                OPCItemQuality(ClientHandles(i)) = OPCquality(Qualities(i))
                OPCItemTimeStamp(ClientHandles(i)) = TimeStamps(i)
            Next i
            OPCdataChanged()
        Catch ex As Exception
            ' Error handling
            If _SilentMode = False Then
                MessageBox.Show("OPC DataChange failed with exception: " + ex.Message, "myOPC Exception", MessageBoxButtons.OK)
            End If
        End Try
    End Sub

    Protected Overridable Sub OPCdataChanged()
        ' Add your crap here
    End Sub
    ''' <summary>
    ''' Writes Data to the PLC tag pointed to by the parameters.  
    ''' Returns TRUE if the write process worked.
    ''' </summary>
    ''' <param name="NewItemValue"></param>
    ''' <param name="ItemClientHandle"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function OPCwrite(ByVal NewItemValue As Object, ByVal ItemClientHandle As Integer) As Boolean
        Try
            ' Check to make sure the imported handle is within the boundaries
            ' between 1 and the number of items.
            If ItemClientHandle > _NumItems Or ItemClientHandle < 1 Then
                ' Error handling
                If _SilentMode = False Then
                    MessageBox.Show("OPC Item Handle Does Not Exist: ", "myOPC Item Missing", MessageBoxButtons.OK)
                End If
                Return False
            End If
            ' Add the imported NewItemValue into the Sync one.
            _SyncItemValues(1) = NewItemValue
            ' Add the imported item handle into the Sync one.
            _SyncItemServerHandles(1) = _ItemServerHandles(ItemClientHandle)
            ' Write the new item value into the requested item.
            _ConnectedGroup.SyncWrite(1, _SyncItemServerHandles, _SyncItemValues, _SyncItemServerErrors)
            If _SyncItemServerErrors(1) <> 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            ' Error handling
            If _SilentMode = False Then
                MessageBox.Show("OPC Write failed with exception: " + ex.Message, "myOPC Exception", MessageBoxButtons.OK)
            End If
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Writes Data to the PLC tag pointed to by the parameters.  
    ''' Returns TRUE if the write process worked.
    ''' </summary>
    ''' <param name="NewItemValue"></param>
    ''' <param name="ItemName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function OPCwrite(ByVal NewItemValue As Object, ByVal ItemName As String) As Boolean
        Dim ItemClientHandle As Integer
        Dim ItemNameFault As Boolean = False
        Try
            For j As Integer = 1 To _NumItems
                If OPCItemNames(j) = ItemName Then
                    ' If Item name is found record the current handle
                    ' so the correct item gets wrote to.
                    ItemClientHandle = j
                    Exit For
                End If
                If j = _NumItems Then
                    ' Error handling
                    If _SilentMode = False Then
                        MessageBox.Show("OPC Item Name Does Not Exist: ", "myOPC Item Missing", MessageBoxButtons.OK)
                    End If
                    ItemNameFault = True
                End If
            Next
            ' Only write value if a valid name was found.
            If ItemNameFault = False Then
                ' Add the imported NewItemValue into the Sync one.
                _SyncItemValues(1) = NewItemValue
                ' Add the found item handle in the Sync one.
                _SyncItemServerHandles(1) = _ItemServerHandles(ItemClientHandle)
                ' Write the new item value into the requested item.
                _ConnectedGroup.SyncWrite(1, _SyncItemServerHandles, _SyncItemValues, _SyncItemServerErrors)
            End If
            If _SyncItemServerErrors(1) <> 0 Or ItemNameFault = True Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            ' Error handling
            If _SilentMode = False Then
                MessageBox.Show("OPC Write failed with exception: " + ex.Message, "myOPC Exception", MessageBoxButtons.OK)
            End If
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Synchronously read data from the PLC tag pointed to by the parameters.  
    ''' Returns True if the read process worked.
    ''' </summary>
    ''' <param name="ItemClientHandle"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function OPCsynchRead(ByVal ItemClientHandle As Integer) As Boolean
        Try
            If ItemClientHandle > _NumItems Or ItemClientHandle < 1 Then
                ' Error handling
                If _SilentMode = False Then
                    MessageBox.Show("OPC Item Handle Does Not Exist: ", "myOPC Item Missing", MessageBoxButtons.OK)
                End If
                Return False
            End If
            ' Add the found item handle in the Sync one.
            _SyncItemServerHandles(1) = _ItemServerHandles(ItemClientHandle)
            ' Read the new item value into the requested item.
            _ConnectedGroup.SyncRead(OPCDataSource.OPCDevice, 1, _SyncItemServerHandles, _SyncItemValues, _SyncItemServerErrors, _SyncItemQuality, _SyncItemTimeStamp)
            If _SyncItemServerErrors(1) = 0 Then
                OPCItemValues(ItemClientHandle) = _SyncItemValues(1)
                OPCItemQuality(ItemClientHandle) = OPCquality(_SyncItemQuality(1))
                OPCItemTimeStamp(ItemClientHandle) = _SyncItemTimeStamp(1)
                Return True
            Else
                OPCItemValues(ItemClientHandle) = Nothing
                OPCItemQuality(ItemClientHandle) = "Bad"
                ' Error handling
                If _SilentMode = False Then
                    MessageBox.Show("OPC read failed: " & _SyncItemServerErrors(1), "myOPC Read Failed", MessageBoxButtons.OK)
                End If
                Return False
            End If
        Catch ex As Exception
            ' Error handling
            If _SilentMode = False Then
                MessageBox.Show("OPC read failed: " & ex.Message, "myOPC Read Failed", MessageBoxButtons.OK)
            End If
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Synchronously read data from the PLC tag pointed to by the parameters.  
    ''' Returns True if the read process worked.
    ''' </summary>
    ''' <param name="ItemName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function OPCsynchRead(ByVal ItemName As String) As Boolean
        Try
            Dim ItemClientHandle As Integer
            For j As Integer = 1 To _NumItems
                If OPCItemNames(j) = ItemName Then
                    ' If Item name is found record the current handle
                    ' so the correct item gets wrote to.
                    ItemClientHandle = j
                    Exit For
                End If
                If j = _NumItems Then
                    ' Error handling
                    If _SilentMode = False Then
                        MessageBox.Show("OPC Item Name Does Not Exist: ", "myOPC Item Missing", MessageBoxButtons.OK)
                    End If
                    Return False
                End If
            Next
            ' Add the found item handle in the Sync one.
            _SyncItemServerHandles(1) = _ItemServerHandles(ItemClientHandle)
            ' Read the new item value into the requested item.
            _ConnectedGroup.SyncRead(OPCDataSource.OPCDevice, 1, _SyncItemServerHandles, _SyncItemValues, _SyncItemServerErrors, _SyncItemQuality, _SyncItemTimeStamp)
            If _SyncItemServerErrors(1) = 0 Then
                OPCItemQuality(ItemClientHandle) = OPCquality(_SyncItemQuality(1))
                OPCItemValues(ItemClientHandle) = _SyncItemValues(1)
                OPCItemTimeStamp(ItemClientHandle) = _SyncItemTimeStamp(1)
                Return True
            Else
                OPCItemValues(ItemClientHandle) = Nothing
                OPCItemQuality(ItemClientHandle) = "Bad"
                ' Error handling
                If _SilentMode = False Then
                    MessageBox.Show("OPC read failed: " & _SyncItemServerErrors(1), "myOPC Read Failed", MessageBoxButtons.OK)
                End If
                Return False
            End If
        Catch ex As Exception
            ' Error handling
            If _SilentMode = False Then
                MessageBox.Show("OPC read failed: " & ex.Message, "myOPC Read Failed", MessageBoxButtons.OK)
            End If
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Synchronously reads all the data in the group from the PLC tag pointed to by the parameters.  
    ''' Returns True if the read process worked.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function OPCsynchReadGroup() As Boolean
        Dim ServerErrorOccured As Boolean = False
        Try
            For j As Integer = 1 To _NumItems
                ' Add the found item handle in the Sync one.
                _SyncItemServerHandles(j) = _ItemServerHandles(j)
            Next
            ' Read the new item value into the requested item.
            _ConnectedGroup.SyncRead(OPCDataSource.OPCDevice, _NumItems, _SyncItemServerHandles, _SyncItemValues, _SyncItemServerErrors, _SyncItemQuality, _SyncItemTimeStamp)
            For k As Integer = 1 To _NumItems
                If _SyncItemServerErrors(k) = 0 Then
                    OPCItemQuality(k) = OPCquality(_SyncItemQuality(k))
                    OPCItemValues(k) = _SyncItemValues(k)
                    OPCItemTimeStamp(k) = _SyncItemTimeStamp(k)
                Else
                    ServerErrorOccured = True
                    OPCItemValues(k) = Nothing
                    OPCItemQuality(k) = "Bad"
                End If
            Next
            If ServerErrorOccured = True Then
                ServerErrorOccured = False
                If _SilentMode = False Then
                    MessageBox.Show("OPC read failed: " & _SyncItemServerErrors(1), "myOPC Read Failed", MessageBoxButtons.OK)
                End If
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            ' Error handling
            If _SilentMode = False Then
                MessageBox.Show("OPC read failed: " & ex.Message, "myOPC Read Failed", MessageBoxButtons.OK)
            End If
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Allows for the OPC connection to be safely disconnected.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Close()
        Try
            ' Remove the items from the group
            _ConnectedGroup.OPCItems.Remove(_ItemCount, _RemoveItemServerHandles, _RemoveItemServerErrors)
            ' Remove the group from the server before closing the form
            _ConnectedOPCServer.OPCGroups.Remove(_GroupName)
            ' Disconnect from OPC server before closing the form
            _ConnectedOPCServer.Disconnect()
        Catch ex As Exception
            ' Error handling
            If _SilentMode = False Then
                MessageBox.Show("OPC Dispose failed with exception: " + ex.Message, "myOPC Exception", MessageBoxButtons.OK)
            End If
        End Try
    End Sub
#End Region

End Class
