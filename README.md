# myOPC
<html><head>
<link rel="stylesheet" type="text/css" href="https://codeproject.freetls.fastly.net/App_Themes/CodeProject/Css/Main.css?dt=2.8.20210930.1">
<base href="https://www.codeproject.com/KB/IP/">
</head>
<body>
<!--
HTML for article "DA 1 OPC Wrapper DLL and Client Example" by Kurt Swaim

Copyright 2012 by Kurt Swaim

<hr class="Divider subdue">
<div>

<!-- Start Article -->
<span id="ArticleContent">
  
<h2>Introduction</h2>
    
<p>This article describes a DLL assembly (<em>myOPC.DLL</em>) that wraps <em>Interop.OPCAutomation.DLL</em> which Interops the COM 
<em>OPCDAAuto.DLL</em>. Both of 
these were created by <a href="http://www.opcfoundation.org/">OPC Foundation</a> and uses DA1 (Data Access 1) which was the original OPC (OLE for Process 
Control) standard. <em>myOPC.DLL</em> makes it easier to create a client that communicates to a device through an OPC server. This article will first describe 
the <em>myOPC.DLL</em> wrapper and then give real world examples of how to use it in a client .NET program.</p>
    
<h2>Background</h2>
    
<p>I was tasked to create a PC program that could interface with a PLC 
(Programmable Logic Controller). At the time I was using VB.NET for my PC 
application development. Searching for a .NET solution I soon discovered 
that OPC was the optimum choice. I'm not going 
into the basics of OPC as this has been done by several people with my 
favorite <a href="http://www.opcconnect.com/">here</a>. To communicate 
with a device (e.g. PLC) through a channel (e.g. Ethernet, Serial) to a PC 
program an OPC server is used. Most modern OPC servers allow for the use 
of three main OPC standards. These are Data Access (DA) version 1-3, 
Unified Architecture (UA), and OPC .NET 3.0. The last one seems to be a 
standard the OPC Foundation would prefer never existed. The UA standard is 
the latest standard and is independent of COM and thus Microsoft. However, 
when looking for example client code that communicated with OPC servers I could 
only find a DA1 example. My organization does not have the cash it takes 
to be an OPC Foundation member where one can find SDK's and examples for the UA 
standard. I found VB.NET client examples on the OPC server 
company 
<a href="http://www.kepware.com/">Kepware</a> web site. With this example 
and my knowledge of object oriented programming I created a wrapper that 
extracted out the ugly details of using the <em>Interop.OPCAutomation.DLL</em>.</p>

<h2>Describing The DLL Code Starting With The StartOPC Function</h2>

<p>The first procedure examined is the function <code>StartOPC()</code>. After creating 
an instance of the <em>myOPC.DLL</em> in the client and setting some properties this 
function starts the OPC connection. The returned Boolean indicates success 
(<code>True</code>) or failure (<code>False</code>) at starting this connection. The below is the 
complete code for this function that I present for clarity. This will be 
followed with detail explanations of sections of this function to describe the more important concepts in this 
function. Please note that the function starts out by checking that 
pivotal properties and fields have been set by the object user. The 
connection will not attempt to start if any of these are empty and will return a 
false.</p>

<pre class="lang-vb" data-language="vbnet" data-allowshrink="True" datacollapse="False">''' <span class="code-summarycomment">&lt;summary&gt;
</span>''' Creates the OPC connection, Groups, Items then activates and reads item values 
''' as long as no errors occur along the way.
''' <span class="code-summarycomment">&lt;/summary&gt;
</span>Public Function StartOPC() As Boolean
    Try
        ' First verify that all of the parameters have been set up
        ' before starting OPC connection.
        If Not _DeadBand = Nothing _
        Or Not _UpdateRate = Nothing _
        Or Not _GroupName = Nothing _
        Or Not _NumItems = Nothing _
        Or ((Not _TopicName = Nothing) Or (Not _CameraName = Nothing) Or _
                  ((Not _ChannelName = Nothing) And (Not _DeviceName = Nothing))) Then
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
            ' Add the items
            _ConnectedGroup.OPCItems.AddItems(_ItemCount, _OPCItemIDs, _
                       _ClientHandles, _ItemServerHandles, _AddItemServerErrors)
            ' For the items added without errors make them active 
            For i = 1 To _NumItems
                If _AddItemServerErrors(i) &lt;&gt; 0 Then
                    ' Set the active type desired.
                    _ActiveState = False
                    ' Get the Servers handle for the desired item. The server handles
                    ' were returned in add item subroutine.
                    _ActiveItemServerHandles(1) = _ItemServerHandles(i)
                    ' Invoke the SetActive operation on the OPC item collection interface
                    _ConnectedGroup.OPCItems.SetActive(1, _ActiveItemServerHandles, _ActiveState, _ActiveItemErrors)
                    ' Inform user of class is not active because of faults.
                    OPCItemActive(i) = False
                Else
                    ' Set the active type desired.
                    _ActiveState = True
                    ' Get the Servers handle for the desired item. The server handles
                    ' were returned in add item subroutine.
                    _ActiveItemServerHandles(1) = _ItemServerHandles(i)
                    ' Invoke the SetActive operation on the OPC item collection interface
                    _ConnectedGroup.OPCItems.SetActive(1, _ActiveItemServerHandles, _ActiveState, _ActiveItemErrors)
                    ' If an error occurred during activation then deactivate the user info.
                    OPCItemActive(i) = True
                End If
                ' Get the Servers handle for the desired item. The server handles were
                ' returned in add item subroutine.
                _SyncItemServerHandles(i) = _ItemServerHandles(i)
                _RemoveItemServerHandles(i) = _ItemServerHandles(i)
            Next
            ' Invoke the SyncRead operation. Remember this call will wait until
            ' completion. The source flag in this case, 'OPCDevice' , is set to
            ' read from device which may take some time.
            _ConnectedGroup.SyncRead(OPCDataSource.OPCDevice, _ItemCount, _
                     _SyncItemServerHandles, _SyncItemValues, _SyncItemServerErrors)
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
                MessageBox.Show("Set the OPC parameters before opening connection", _
                                 "Error", MessageBoxButtons.OK)
            End If
            Return False
        End If
    Catch ex As Exception
        ' Error handling
            _ConnectedOPCServer = Nothing
        If _SilentMode = False Then
            MessageBox.Show("OPC server connect failed with exception: " + _
                       ex.Message, "myOPC Exception", MessageBoxButtons.OK)
        End If
        Return False
    End Try
End Function</pre>
<h4>Creating The OPC Server Object</h4>

<p>The first thing required to start the OPC is to instantiate a new OPC server 
object. Once this object is created a connection is made to the 
actual 
OPC server. The first parameter, <code>_OPCserver</code>, is a string of the name of 
the OPC server being used called ProgID. These values were discovered 
using the discovery example found on the Kepwareweb site. The second 
parameter is called <code>Node</code> and is an <code>object</code>. It is not clear to me what this 
is used for and because the connection worked without it I just used what the Kepware example used (i.e. ""). The way that user of this DLL specifies 
the OPC server ProgID is through the <code>OPCServerType</code> integer property which must 
be set before starting the OPC server. The property in turns sets the 
<code>_OPCserver</code> to the proper string. If you use a different OPC server company 
than I do you will have to modify the DLL for a 4th OPC type that sets the 
appropriate string for your OPC server.</p>
<pre data-language="text" data-allowshrink="True" datacollapse="False">OPCServerType Value&nbsp;&nbsp;&nbsp;=&gt;     _OPCserver Value
1      =&gt;     "RSLinx OPC Server"
2      =&gt;     "Kepware.KEPServerEX.V5"
3      =&gt;     "Cognex In-Sight OPC Server"</pre>
<pre class="lang-vb" data-language="vbnet" data-allowshrink="True" datacollapse="False">' Create a new OPC Server object
_ConnectedOPCServer = New OPCServer
' Attempt to connect with the server
_ConnectedOPCServer.Connect(_OPCserver, "")</pre>
<h4>Creating The OPC Group</h4>
<p>The OPC group is next to be created and defined. The group houses the OPC items and define some of their behavior. The first step is to make 
groups default to be active. Next the dead band is set via the property <code>DeadBand</code>. The dead band is an integer that is kept to a value of 0 to 99 
via the <code>DeadBand</code> property. The dead band is only valid for an asynch group. If set to zero then any change in the group items will cause an 
asynch read of the changed item. If the value is 10, for example, then the 
asynch read will occur when the item value changes +/- 10%. Next the group 
is created with the string name of the group set with <code>_GroupName</code>. I found that this 
string is useful in troubleshooting. The update rate is next to be 
defined with the property <code>UpdateRate</code>. This property is an integer and 
represents the update rate of the group in milliseconds. When I have 
multiple connections to the same device I set each update rates at a different 
prime number. Also keep in mind it is possible for some OPC servers to 
optionally over ride this rate. The <code>UpdateRate</code> is followed with setting 
the Boolean property called <code>AsynchMode</code>. When <code>AsynchMode</code> is set to False 
the values of the items in the group can only be discovered through a manual 
client OPC read. Inversely when set to True a method will be raised every 
time the value changes based on <code>DeadBand</code> and <code>UpdateRate</code>. Synchronous and 
Asynchronous modes will be described more later in this article. The last 
group thing modified is the <code>DefaultIsActive</code> set to False forcing the items to be 
added in a non active state. Items should not be made active until they 
are confirmed working with no faults.</p>

<pre class="lang-vb" data-language="vbnet" data-allowshrink="True" datacollapse="False">' Set the desire active state for the group
_ConnectedOPCServer.OPCGroups.DefaultGroupIsActive = True
' Set the desired percent dead band
_ConnectedOPCServer.OPCGroups.DefaultGroupDeadband = _DeadBand
' Add the group
_ConnectedGroup = _ConnectedOPCServer.OPCGroups.Add(_GroupName)
' Set the update rate for the group
_ConnectedGroup.UpdateRate = _UpdateRate
' Mark this group to receive asynchronous updates via the DataChange event.
_ConnectedGroup.IsSubscribd = _AsynchMode
' Setting the '.DefaultIsActive' property forces all items we are about to
' add to the group to be added in a non active state.
_ConnectedGroup.OPCItems.DefaultIsActive = False</pre>
<h4>Creating The OPC Items</h4>

<p>
After the OPC group has been created and defined the OPC items can be added. 
This process uses the <code>NumItems</code> property to set up the <code>For</code> loop.&nbsp;
<code>NumItems</code> 
also re-DIMs the Item arrays. Make sure the number of itemID's added 
matches the <code>NumItems</code>. The two tricky items for me creating this DLL was 
determining the progID and the ItemID. From the below code it can be seen 
that the string for ItemID is formatted significantly differently for each OPC 
server type. The only common portion of this format is the <code>OPCItemNames</code> 
array defined by the client. This array gives the address in the device of the items 
being tracked, I call them tags. The format 
of these tags can also be confusing. Fortunately Kepware has done a great 
job of documenting them. My example will use Allen Bradley PLC tags. Notice that 
certain properties pertain to the OPC server type chosen. For example if 
you are using a RSLinx OPC server only the <code>TopicName</code> is required and the 
<code>ChannelName</code>, <code>DeviceName</code>, and <code>CameraName</code> can be left undefined. Note two 
things from the below For loop, the first being that the arrays used in an OPC 
server is One based not Zero. Second notice inside the For loop the 
<code>_ClientHandles</code> is defined with the For loop index. Once the ItemID array has 
been defined then all of the items set up by the client is added at once. 
This method calls to the connected group to add items and has five parameters. 
The <code>_ItemCount</code> is equal to <code>_NumItems</code>. I know what your thinking and no I 
don't know why I didn't just use <code>_NumItems</code>. The array <code>_OPCItemIDs</code> is what 
we just created which contains the OPC items ID (i.e. the complete address). The 
<code>_ClientHandles</code>, also 
just created, represents a number used by the client for reading and writing the 
specific item desired. The forth and fifth parameters are set by the OPC 
server call not the client. The <code>_ItemServerHandles</code> is similar to the 
<code>_ClientHandles</code> except that it is created by the server. </p><pre data-allowshrink="True" datacollapse="False">_ItemServerHandles</pre> will be used for synchronous reads. The last parameter 
is used to inform of any errors that occurred.<p></p>

<pre class="lang-vb" data-language="vbnet" data-allowshrink="True" datacollapse="False">' Assemble the OPC Item IDs
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
' Add the items
_ConnectedGroup.OPCItems.AddItems(_ItemCount, _OPCItemIDs, _
        _ClientHandles, _ItemServerHandles, _AddItemServerErrors)</pre>
<h4>Activate The OPC Items</h4>
<p>Now that the items have been added to the group the items without errors are made active. I know this bit of code could use some refactoring. 
After making the items active all of the items in the group are read. I have since wondered if I should have added 
this functionality. The problem arrives in asynchronous mode in that the asynch method is called when the <code>StartOPC()</code> function is ran. Depending on 
your application this can be a problem. The final thing I want to mention here is about the <code>SilentMode</code> Boolean property. Looking at the full code 
above it can be seen that all of this logic is wrapped in a try..catch block. This message box is used to inform the client that an error occurred when the 
<code>SilentMode</code> is set to False. Once I go from development to production I set <code>SilentMode</code> 
property to True as I have worked out any bugs.</p>
<pre class="lang-vb" data-language="vbnet" data-allowshrink="True" datacollapse="False">' For the items added without errors make them active 
For i = 1 To _NumItems
    If _AddItemServerErrors(i) &lt;&gt; 0 Then
        ' Set the active type desired.
        _ActiveState = False
        ' Get the Servers handle for the desired item. The server handles
        ' were returned in add item subroutine.
        _ActiveItemServerHandles(1) = _ItemServerHandles(i)
        ' Invoke the SetActive operation on the OPC item collection interface
        _ConnectedGroup.OPCItems.SetActive(1, _ActiveItemServerHandles, _ActiveState, _ActiveItemErrors)
        ' Inform user of class is not active because of faults.
        OPCItemActive(i) = False
    Else
        ' Set the active type desired.
        _ActiveState = True
        ' Get the Servers handle for the desired item. The server handles
        ' were returned in add item subroutine.
        _ActiveItemServerHandles(1) = _ItemServerHandles(i)
        ' Invoke the SetActive operation on the OPC item collection interface
        _ConnectedGroup.OPCItems.SetActive(1, _ActiveItemServerHandles, _ActiveState, _ActiveItemErrors)
        ' If an error occurred during activation then deactivate the user info.
        OPCItemActive(i) = True
    End If
    ' Get the Servers handle for the desired item. The server handles were
    ' returned in add item subroutine.
    _SyncItemServerHandles(i) = _ItemServerHandles(i)
    _RemoveItemServerHandles(i) = _ItemServerHandles(i)
Next</pre>
<h2>Describing other important methods in the DLL assembly</h2>
<h4>The Asynchronous Functionality</h4>
<p>If the client sets the OPC group to be asynchronous the below method will fire 
every time one or more items changes within the dead band. This method 
receives the number of items effected and their corresponding client handles 
from the OPC server. 
Through the use of this information this method will update the item fields that 
changed related to the OPC item value, quality, and time stamp. Once the 
changed items have been updated the method <code>OPCdataChanged()</code> is called. 
This method is meant for the client to over-ride it so it can be informed of a 
change of data. I will show this in the example but this DLL can be used 
in two main ways. If used in synchronous mode then <em>myOPC.DLL</em> is 
instantiated directly. Otherwise a 
class should be created that inherits from <em>myOPC.DLL</em>. In this way this new class can over-ride the 
<code>OPCdataChanged()</code> 
method and use it as required.</p>

<pre class="lang-vb" data-language="vbnet" data-allowshrink="True" datacollapse="False">Private Sub _ConnectedGroup_DataChange(ByVal TransactionID As Integer, _
        ByVal NumItems As Integer, ByRef ClientHandles As System.Array, _
        ByRef ItemValues As System.Array, ByRef Qualities As System.Array, _
        ByRef TimeStamps As System.Array) Handles _ConnectedGroup.DataChange
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
                MessageBox.Show("OPC DataChange failed with exception: " + _
                          ex.Message,"myOPC Exception",MessageBoxButtons.OK)
            End If
    End Try
End Sub

Protected Overridable Sub OPCdataChanged()
    ' Add your crap here
End Sub</pre>

<h4>Synchronous Item Read</h4>

<p>
If the client sets the OPC group to be Synchronous the below Boolean function 
can be used to manually read individual items. If all goes well a True will be 
returned and the corresponding field value will be updated (i.e., <code>OPCItemValues(ItemClientHandle)</code>). 
<code>OPCItemValues()</code> is an array of objects 
that contains the OPC item value regardless of the mode (e.g., synchronous or 
asynchronous). The <code>ItemClientHandle</code> is the parameter required for the 
<code>OPCsynchRead()</code> function that points to the desired item to be updated. Notice that because 
<code>OPCItemValues()</code> are objects a cast will be required before the client can 
effectively use them. The first thing this function does is verify that 
the value of <code>ItemClientHandle</code> is legitimate. After looking up the server 
item handle this code performs a <code>SynchRead</code> on the specified item. If the 
synchronous read returns no errors then the specified OPC item value, quality, 
and time stamp will be updated. Note that a nearly identical function also 
exists that uses the item name as the parameter instead of the item handle.</p>

<pre class="lang-vb" data-language="vbnet" data-allowshrink="True" datacollapse="False">''' <span class="code-summarycomment">&lt;summary&gt;;
</span>''' Synchronously read data from the PLC tag pointed to by the parameters.
''' Returns True if the read process worked.
''' <span class="code-summarycomment">&lt;/summary&gt;
</span>Public Function OPCsynchRead(ByVal ItemClientHandle As Integer) As Boolean
    Try
        If ItemClientHandle &gt; _NumItems Or ItemClientHandle &lt; 1  Then
            ' Error handling
            If _SilentMode = False Then
                MessageBox.Show("OPC Item Handle Does Not Exist: ", _
                          "myOPC Item Missing", MessageBoxButtons.OK)
            End If
            Return False
        End If
        ' Add the found item handle in the Sync one.
        _SyncItemServerHandles(1) = _ItemServerHandles(ItemClientHandle)
        ' Read the new item value into the requested 
        _ConnectedGroup.SyncRead(OPCDataSource.OPCDevice, 1, _SyncItemServerHandles, _
                _SyncItemValues, _SyncItemServerErrors, _SyncItemQuality, _SyncItemTimeStamp)
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
                MessageBox.Show("OPC read failed: " &amp; _SyncItemServerErrors(1), _
                             "myOPC Read Failed", MessageBoxButtons.OK)
            End If
            Return False
        End If
    Catch ex As Exception
        ' Error handling
        If _SilentMode = False Then
            MessageBox.Show("OPC read failed: " &amp; ex.Message, _
                        "myOPC Read Failed", MessageBoxButtons.OK)
        End If
        Return False
    End Try
End Function</pre>
<h4>Synchronous Group Read</h4>
<p>Alternatively we can read all items within the group at once with the below Boolean function. It requires no parameters and returns a True if all goes well. 
The logic is nearly the same to the above except that it loops through all the item handles and reads them.</p>
<pre class="lang-vb" data-language="vbnet" data-allowshrink="True" datacollapse="False">''' <span class="code-summarycomment">&lt;summary&gt;
</span>''' Synchronously reads all the data in the group from the PLC tag pointed to by the parameters.
''' Returns True if the read process worked.
''' <span class="code-summarycomment">&lt;/summary&gt;
</span>Public Function OPCsynchReadGroup() As Boolean
    Dim ServerErrorOccured As Boolean = False
    Try
        For j As Integer = 1 To _NumItems
        ' Add the found item handle in the Sync one.
            _SyncItemServerHandles(j) = _ItemServerHandles(j)
        Next
        ' Read the new item value into the requested item.
        _ConnectedGroup.SyncRead(OPCDataSource.OPCDevice, _NumItems, _SyncItemServerHandles, _
                  _SyncItemValues, _SyncItemServerErrors, _SyncItemQuality, _SyncItemTimeStamp)
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
                MessageBox.Show("OPC read failed: " &amp; _SyncItemServerErrors(1), _
                            "myOPC Read Failed", MessageBoxButtons.OK)
            End If
            Return False
        Else
            Return True
        End If
    Catch ex As Exception
        ' Error handling
        If _SilentMode = False Then
            MessageBox.Show("OPC read failed: " &amp; ex.Message, _
                          "myOPC Read Failed", MessageBoxButtons.OK)
        End If
        Return False
    End Try
End Function</pre>
<h4>Synchronous Item Write</h4>

<p>
If an item needs to be changed by your client program than <code>OPCwrite()</code> is the 
Boolean function for you. Again if all goes well this function will return 
a <code>True</code>. The function accepts two parameters the first representing the new 
value and the second the handle of the item. Once the handle is verified 
the function writes the items new value. If any errors occur along the 
line a False is returned. Note that a nearly identical function also exists that 
uses the item name as the parameter instead of the item handle.</p>

<pre class="lang-vb" data-language="vbnet" data-allowshrink="True" datacollapse="False">''' <span class="code-summarycomment">&lt;summary&gt;
</span>''' Writes Data to the PLC tag pointed to by the parameters.
''' Returns TRUE if the write process worked.
''' <span class="code-summarycomment">&lt;/summary&gt;
</span>Public Function OPCwrite(ByVal NewItemValue As Object,ByVal ItemClientHandle As Integer) As Boolean
    Try
        ' Check to make sure the imported handle is within the boundaries
        ' between 1 and the number of items.
        If ItemClientHandle &gt; _NumItems Or ItemClientHandle &lt; 1 Then
            ' Error handling
            If _SilentMode =False Then
                MessageBox.Show("OPC Item Handle Does Not Exist: ", _
                          "myOPC Item Missing", MessageBoxButtons.OK)
            End If
            Return False
        End If
        ' Add the imported NewItemValue into the Sync one.
        _SyncItemValues(1) = NewItemValue
        ' Add the imported item handle into the Sync
        _SyncItemServerHandles(1) = _ItemServerHandles(ItemClientHandle)
        ' Write the new item value into the requested
        _ConnectedGroup.SyncWrite(1, _SyncItemServerHandles, _SyncItemValues, _SyncItemServerErrors)
        If _SyncItemServerErrors(1) &lt;&gt; 0Then
            Return False
        Else
            Return True
        End If
    Catch ex As Exception
        ' Error handling
        If _SilentMode = False Then
            MessageBox.Show("OPC Write failed with exception: " + _
                      ex.Message, "myOPC Exception", MessageBoxButtons.OK)
        End If
        Return False
    End Try
End Function</pre>

<h4>Close Connection</h4>

<p>Be a good steward and close the OPC connection with the below method. If 
you fail to close the connection while closing the program and/or form the group 
and its items will remain active at the OPC server. This method closes in 
the same way it opens but in reverse order.</p>

<pre class="lang-vb" data-language="vbnet" data-allowshrink="True" datacollapse="False">''' <span class="code-summarycomment">&lt;summary&gt;
</span>''' Allows for the OPC connection to be safely disconnected.
''' <span class="code-summarycomment">&lt;/summary&gt;
</span>Public Sub Close()
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
            MessageBox.Show("OPC Dispose failed with exception: " + _
                          ex.Message,"myOPC Exception",MessageBoxButtons.OK)
        End If
    End Try
End Sub</pre>

<h2>Download Files</h2>

<p>Two main zip files are provided. The first is visual studio 2010 solution that compiles the 
<em>myOPC.dll</em> assembly. If you wish to improve or add to this object then you will need this zip file. 
If your just going to use the DLL file for your client needs then the second zip file is all that is required. 
Inside this zip file you will find several files. The files <em>myOPC.dll</em> and 
<em>Interop.OPCAutomation.dll</em> both should be referenced by your client file and 
their properties "Copy Local" set to true. The file 
<em>Interop.OPCAutomation.dll</em> will have to have its property "Embed Interop Types" 
set to False 
first. On every machine your client program uses <em>myOPC.dll</em> the 
assembly <em>OPCDAAuto.dll</em> must also be located and registered to the that machines operating 
system. This registration is done via the included <em>update_OPCdaauto.bat</em> 
file. If you don't do this (and you will forget) the assembly <em>myOPC.dll</em> 
will not function. I usually copy both <em>OPCDAAuto.dll</em> and the batch file to 
the System32 directory of the machine I'm going to use than run the batch file.</p>

<h2>Client Example</h2>

<p>The below example is a simple windows form that creates a synchronous and asynchronous OPC connection when the form loads. Don't forget to reference 
<em>myOPC.dll</em> and <em>Interop.OPCAutomation.dll</em> 
and register <em>OPCDAAuto.dll</em> just explained. The below is a common scenario 
for my use of an OPC connection to a PLC (Allen Bradley in this case). 
The example also uses a Kepware server already set up with <code>ChannelName</code> and 
<code>DeviceName</code>. When something in a manufacturing process occurs like the end of a machine cycle a bit 
usually gets set indicating such. This bit is being watched by an 
asynchronous group and runs the inherited/overloaded method <code>OPCdataChanged()</code>. Then code 
is added to <code>OPCdataChanged</code> to 
grab and record process data and reset the trigger flag. The last thing I 
want you to remember is the need to convert OPC items before using them as they 
come back as objects.</p>

<pre class="lang-vb" data-language="vbnet" data-allowshrink="True" datacollapse="False">Public Class Form1
    ' Create OPC instances
    Public myAsynchOpc As AsynchOpc = New AsynchOpc
    Public mySynchOpc As myOPC.myOPC = New myOPC.myOPC

    Private Sub Form1_Load(sender As Object, e As EventArgs) HandlesMyBase.Load
        ' Define Asynchronous Parameters
        myAsynchOpc.AsynchMode = True
        myAsynchOpc.ChannelName = "my2Channel"
        myAsynchOpc.DeviceName = "my2Device"
        myAsynchOpc.DeadBand = 0
        myAsynchOpc.GroupName = "myAsynchGroup"
        myAsynchOpc.NumItems = 1
        myAsynchOpc.OPCServerType = 2
        myAsynchOpc.SilentMode = False
        myAsynchOpc.UpdateRate = 100
        myAsynchOpc.OPCItemNames(1) = "myTrigger"

        ' Define Synchronous Parameters
        mySynchOpc.AsynchMode = False
        mySynchOpc.ChannelName = "my2Channel"
        mySynchOpc.DeviceName = "my2Device"
        mySynchOpc.DeadBand = 0
        mySynchOpc.GroupName = "mySynchGroup"
        mySynchOpc.NumItems = 2
        mySynchOpc.OPCServerType = 2
        mySynchOpc.SilentMode = False
        mySynchOpc.UpdateRate = 100
        mySynchOpc.OPCItemNames(1) = "myData[0]"
        mySynchOpc.OPCItemNames(2) = "myData[1]"

        Try 'Start OPC's
            Dim asynchOpcStarted As Boolean = myAsynchOpc.StartOPC()
            Dim synchOpcStarted As Boolean = mySynchOpc.StartOPC()
            ' Add logic to verify they started
        Catch ex As Exception
            ' Error Logic Here
        End Try
    End Sub

    ' Don't forget to be a good steward and close OPC connections
    Private Sub Form1_FormClosing(sender As Object, _
                e As FormClosingEventArgs)Handles MyBase.FormClosing
        myAsynchOpc.Close()
        mySynchOpc.Close()
    End Sub

    Public Class AsynchOpc
            Inherits myOPC.myOPC

        Protected Overrides Sub OPCdataChanged()
            MyBase.OPCdataChanged()

            ' Shown for clarity could inline this
            Dim theTrigger As Boolean = Convert.ToBoolean(OPCItemValues(1))

            If theTrigger = True Then
                ' If the trigger is on read data
                Form1.mySynchOpc.OPCsynchReadGroup()
                ' Grab Data
                Dim Data0 As Int32 = Convert.ToInt32(Form1.mySynchOpc.OPCItemValues(1))
                Dim Data1 As Int32 = Convert.ToInt32(Form1.mySynchOpc.OPCItemValues(2))
                ' Reset Trigger
                OPCwrite(0,1)
            End If
        End Sub
    End Class
End Class</pre>
<h2>Final Thoughts</h2>
<p>This assembly was my first OO attempt. I had moved from VB6 to VB.NET with little knowledge of the OO world and its best practices. I made some 
major mistakes like using a message box instead of raising an exception. I also made minor mistakes like incorrect casing and things in between. So 
why didn't I change it?&nbsp; I did not have time and the assembly worked well for what I was doing at the time. I am currently creating a new version of 
this assembly with C# and <a href="http://www.advosol.com/">Advosol</a> wrapper instead of OPC foundation. This will give me DA3 instead of DA1 and with 
my improved .NET knowledge should result in a more efficient and powerful way of communicating with OPC devices.</p>
<h2>History</h2>
<ul>
<li>5/2008 - <em>myOPC.dll</em> created.</li><li>6/2010 - <em>myOPC.dll</em> final revision.</li><li>11/2012 - Article submitted.</li></ul>


</span>
<!-- End Article -->




</div> 


</body></html>
