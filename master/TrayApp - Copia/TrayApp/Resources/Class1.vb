Imports System.Runtime.InteropServices
Module modSiemens

    'Modulo di Gestione PLC Siemens V0.1
    '
    'E' stato eseguito un Porting dalla VB6 corrette solo alcune funzioni che andavo in crash dato che i 
    'vecchi tipi Long in VB .NET sono Integer.
    '
    '
    'Su sistema a 64bit non funziona la DLL, molto probabilmente la dll e' stata scritta per i 32 bit
    'da un errore di tipo System.BadImageFormatException
    '
    'Sfruttata la DLL vecchio stampo, dato che ho abilitato Gli InteropServices, quindi non e' necessario il Wrapper per .net
    '
    'Testato con Cavo PPI Siemens
    'Testato su Windows XP Professional 32Bit
    'Testato su Windows Vista Enterprise 32Bit
    'Velocità Circa 120 mS a lettura

    Public Structure PLC_AdrType
        Dim iConnessione As Integer
        Dim sPLC_IP As String
        Dim sPortaSeriale As String
        Dim sBaudRate As String
        Dim sComParity As String
        Dim sConnessione As String
        Dim iMpiPpi As Integer
        Dim ph As Integer
        Dim dInterf As Integer
        Dim dConn As Integer
        Dim retval As Integer
        Dim sMessaggio As String
        Dim lUsedTime As Integer
        Dim iRack As Integer
        Dim iSlot As Integer
        Dim lCommErr As Integer
        Dim lCommOK As Integer
    End Structure

    ' Numero massimo di byte in un'unico telegramma in lettura
    Private Const br1MaxByteRead = 200
    ' Numero massimo di byte in un'unico telegramma in scrittura
    Private Const br1MaxByteWrite = 200
    '
    '    Protocol types to be used with newInterface:
    '
    Private Const daveProtoMPI = 0      '  MPI for S7 300/400
    Private Const daveProtoMPI2 = 1    '  MPI for S7 300/400, "Andrew's version"
    Private Const daveProtoMPI3 = 2    '  MPI for S7 300/400, Step 7 Version, not yet implemented
    Private Const daveProtoPPI = 10    '  PPI for S7 200
    Private Const daveProtoAS511 = 20    '  S5 via programming interface
    Private Const daveProtoS7online = 50    '  S7 using Siemens libraries & drivers for transport
    Private Const daveProtoISOTCP = 122 '  ISO over TCP
    Private Const daveProtoISOTCP243 = 123 '  ISO o?ver TCP with CP243
    Private Const daveProtoMPI_IBH = 223   '  MPI with IBH NetLink MPI to ethernet gateway */
    Private Const daveProtoPPI_IBH = 224   '  PPI with IBH NetLink PPI to ethernet gateway */
    Private Const daveProtoUserTransport = 255 '  Libnodave will pass the PDUs of S7 Communication to user defined call back functions.
    '
    '    ProfiBus speed constants:
    '
    Private Const daveSpeed9k = 0
    Private Const daveSpeed19k = 1
    Private Const daveSpeed187k = 2
    Private Const daveSpeed500k = 3
    Private Const daveSpeed1500k = 4
    Private Const daveSpeed45k = 5
    Private Const daveSpeed93k = 6
    '
    '    S7 specific constants:
    '
    Private Const daveBlockType_OB = "8"
    Private Const daveBlockType_DB = "A"
    Private Const daveBlockType_SDB = "B"
    Private Const daveBlockType_FC = "C"
    Private Const daveBlockType_SFC = "D"
    Private Const daveBlockType_FB = "E"
    Private Const daveBlockType_SFB = "F"
    '
    ' Use these constants for parameter "area" in daveReadBytes and daveWriteBytes
    '
    Private Const daveSysInfo = &H3      '  System info of 200 family
    Private Const daveSysFlags = &H5   '  System flags of 200 family
    Private Const daveAnaIn = &H6      '  analog inputs of 200 family
    Private Const daveAnaOut = &H7     '  analog outputs of 200 family
    Private Const daveP = &H80          ' direct access to peripheral adresses
    Private Const daveInputs = &H81
    Private Const daveOutputs = &H82
    Private Const daveFlags = &H83
    Private Const daveDB = &H84 '  data blocks
    Private Const daveDI = &H85  '  instance data blocks
    Private Const daveV = &H87      ' don't know what it is
    Private Const daveCounter = 28  ' S7 counters
    Private Const daveTimer = 29    ' S7 timers
    Private Const daveCounter200 = 30       ' IEC counters (200 family)
    Private Const daveTimer200 = 31         ' IEC timers (200 family)
    '
    Private Const daveOrderCodeSize = 21    ' Length of order code (MLFB number)
    '
    '    Library specific:
    '
    '
    '    Result codes. Genarally, 0 means ok,
    '    >0 are results (also errors) reported by the PLC
    '    <0 means error reported by library code.
    '
    Private Const daveResOK = 0                        ' means all ok
    Private Const daveResNoPeripheralAtAddress = 1     ' CPU tells there is no peripheral at address
    Private Const daveResMultipleBitsNotSupported = 6  ' CPU tells it does not support to read a bit block with a
    ' length other than 1 bit.
    Private Const daveResItemNotAvailable200 = 3       ' means a a piece of data is not available in the CPU, e.g.
    ' when trying to read a non existing DB or bit bloc of length<>1
    ' This code seems to be specific to 200 family.
    Private Const daveResItemNotAvailable = 10         ' means a a piece of data is not available in the CPU, e.g.
    ' when trying to read a non existing DB
    Private Const daveAddressOutOfRange = 5            ' means the data address is beyond the CPUs address range
    Private Const daveWriteDataSizeMismatch = 7        ' means the write data size doesn't fit item size
    Private Const daveResCannotEvaluatePDU = -123
    Private Const daveResCPUNoData = -124
    Private Const daveUnknownError = -125
    Private Const daveEmptyResultError = -126
    Private Const daveEmptyResultSetError = -127
    Private Const daveResUnexpectedFunc = -128
    Private Const daveResUnknownDataUnitSize = -129
    Private Const daveResShortPacket = -1024
    Private Const daveResTimeout = -1025
    '
    '    Max number of bytes in a single message.
    '
    Private Const daveMaxRawLen = 2048
    '
    '    Some definitions for debugging:
    '
    Private Const daveDebugRawRead = &H1            ' Show the single bytes received
    Private Const daveDebugSpecialChars = &H2       ' Show when special chars are read
    Private Const daveDebugRawWrite = &H4           ' Show the single bytes written
    Private Const daveDebugListReachables = &H8     ' Show the steps when determine devices in MPI net
    Private Const daveDebugInitAdapter = &H10       ' Show the steps when Initilizing the MPI adapter
    Private Const daveDebugConnect = &H20           ' Show the steps when connecting a PLC
    Private Const daveDebugPacket = &H40
    Private Const daveDebugByte = &H80
    Private Const daveDebugCompare = &H100
    Private Const daveDebugExchange = &H200
    Private Const daveDebugPDU = &H400      ' debug PDU handling
    Private Const daveDebugUpload = &H800   ' debug PDU loading program blocks from PLC
    Private Const daveDebugMPI = &H1000
    Private Const daveDebugPrintErrors = &H2000     ' Print error messages
    Private Const daveDebugPassive = &H4000
    Private Const daveDebugErrorReporting = &H8000
    Private Const daveDebugOpen = &H8000
    Private Const daveDebugAll = &H1FFFF

    Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Integer, ByRef Source As Integer, ByVal Length As Integer)
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)

    '
    '    Set and read debug level:
    '
    Private Declare Sub daveSetDebug Lib "libnodave.dll" (ByVal level As Long)
    Private Declare Function daveGetDebug Lib "libnodave.dll" () As Long
    '
    ' You may wonder what sense it might make to set debug level, as you cannot see
    ' messages when you opened excel or some VB application from Windows GUI.
    ' You can invoke Excel from the console or from a batch file with:
    ' <myPathToExcel>\Excel.Exe <MyPathToXLS-File>VBATest.XLS >ExcelOut
    ' This will start Excel with VBATest.XLS and all debug messages (and a few from Excel itself)
    ' go into the file ExcelOut.
    '
    '    Error code to message string conversion:
    '    Call this function to get an explanation for error codes returned by other functions.
    '
    '
    ' The folowing doesn't work properly. A VB string is something different from a pointer to char:
    '
    ' Private Declare Function daveStrerror Lib "libnodave.dll" Alias "daveStrerror" (ByVal en As Long) As String
    '
    Private Declare Function daveInternalStrerror Lib "libnodave.dll" Alias "daveStrerror" (ByVal en As Integer) As Integer
    ' So, I added another function to libnodave wich copies the text into a VB String.
    ' This function is still not useful without some code araound it, so I call it "internal"
    Private Declare Sub daveStringCopy Lib "libnodave.dll" (ByVal internalPointer As Integer, ByVal s As String)
    '
    ' Setup a new interface structure using a handle to an open port or socket:
    '
    Private Declare Function daveNewInterface Lib "libnodave.dll" (ByVal fd1 As Integer, ByVal fd2 As Integer, ByVal name As String, ByVal localMPI As Integer, ByVal protocol As Integer, ByVal speed As Integer) As Integer
    '
    ' Setup a new connection structure using an initialized daveInterface and PLC's MPI address.
    ' Note: The parameter di must have been obtained from daveNewinterface.
    '
    Private Declare Function daveNewConnection Lib "libnodave.dll" (ByVal di As Integer, ByVal mpi As Integer, ByVal Rack As Integer, ByVal Slot As Integer) As Integer
    '
    '    PDU handling:
    '    PDU is the central structure present in S7 communication.
    '    It is composed of a 10 or 12 byte header,a parameter block and a data block.
    '    When reading or writing values, the data field is itself composed of a data
    '    header followed by payload data
    '
    '    retrieve the answer:
    '    Note: The parameter dc must have been obtained from daveNewConnection.
    '
    Private Declare Function daveGetResponse Lib "libnodave.dll" (ByVal dc As Long) As Long
    '
    '    send PDU to PLC
    '    Note: The parameter dc must have been obtained from daveNewConnection,
    '          The parameter pdu must have been obtained from daveNewPDU.
    '
    Private Declare Function daveSendMessage Lib "libnodave.dll" (ByVal dc As Long, ByVal pdu As Long) As Long
    '******
    '
    'Utilities:
    '
    '****
    '*
    '    Hex dump PDU:
    '
    Private Declare Sub daveDumpPDU Lib "libnodave.dll" (ByVal pdu As Long)
    '
    '    Hex dump. Write the name followed by len bytes written in hex and a newline:
    '
    Private Declare Sub daveDump Lib "libnodave.dll" (ByVal name As String, ByVal pdu As Long, ByVal Length As Long)
    '
    '    names for PLC objects. This is again the intenal function. Use the wrapper code below.
    '
    Private Declare Function daveInternalAreaName Lib "libnodave.dll" Alias "daveAreaName" (ByVal en As Long) As Long
    Private Declare Function daveInternalBlockName Lib "libnodave.dll" Alias "daveBlockName" (ByVal en As Long) As Long
    '
    '   swap functions. They change the byte order, if byte order on the computer differs from
    '   PLC byte order:
    '
    Private Declare Function daveSwapIed_16 Lib "libnodave.dll" (ByVal x As Long) As Long
    Private Declare Function daveSwapIed_32 Lib "libnodave.dll" (ByVal x As Long) As Long
    '
    '    Data conversion convenience functions. The older set has been removed.
    '    Newer conversion routines. As the terms WORD, INT, INTEGER etc have different meanings
    '    for users of different programming languages and compilers, I choose to provide a new
    '    set of conversion routines named according to the bit length of the value used. The 'U'
    '    or 'S' stands for unsigned or signed.
    '
    '
    '    Get a value from the position b points to. B is typically a pointer to a buffer that has
    '    been filled with daveReadBytes:
    '
    Private Declare Function toPLCfloat Lib "libnodave.dll" (ByVal f As Single) As Single
    Private Declare Function daveToPLCfloat Lib "libnodave.dll" (ByVal f As Single) As Long
    '
    ' Copy and convert value of 8,16,or 32 bit, signed or unsigned at position pos
    ' from internal buffer:
    '
    Private Declare Function daveGetS8from Lib "libnodave.dll" (ByRef buffer As Byte) As Long
    Private Declare Function daveGetU8from Lib "libnodave.dll" (ByRef buffer As Byte) As Long
    Private Declare Function daveGetS16from Lib "libnodave.dll" (ByRef buffer As Byte) As Long
    Private Declare Function daveGetU16from Lib "libnodave.dll" (ByRef buffer As Byte) As Long
    Private Declare Function daveGetS32from Lib "libnodave.dll" (ByRef buffer As Byte) As Long
    '
    ' Is there an unsigned long? Or a longer integer than long? This doesn't work.
    ' Private Declare Function daveGetU32from Lib "libnodave.dll" (ByRef buffer As Byte) As Long
    '
    Private Declare Function daveGetFloatfrom Lib "libnodave.dll" (ByRef buffer As Byte) As Single
    '
    ' Copy and convert a value of 8,16,or 32 bit, signed or unsigned from internal buffer. These
    ' functions increment an internal buffer position. This buffer position is set to zero by
    ' daveReadBytes, daveReadBits, daveReadSZL.
    '
    Private Declare Function daveGetS8 Lib "libnodave.dll" (ByVal dc As Long) As Long
    Private Declare Function daveGetU8 Lib "libnodave.dll" (ByVal dc As Integer) As Integer
    Private Declare Function daveGetS16 Lib "libnodave.dll" (ByVal dc As Long) As Long
    Private Declare Function daveGetU16 Lib "libnodave.dll" (ByVal dc As Long) As Long
    Private Declare Function daveGetS32 Lib "libnodave.dll" (ByVal dc As Long) As Long
    '
    ' Is there an unsigned long? Or a longer integer than long? This doesn't work.
    'Private Declare Function daveGetU32 Lib "libnodave.dll" (ByVal dc As Long) As Long
    Private Declare Function daveGetFloat Lib "libnodave.dll" (ByVal dc As Long) As Single
    '
    ' Read a value of 8,16,or 32 bit, signed or unsigned at position pos from internal buffer:
    '
    Private Declare Function daveGetS8At Lib "libnodave.dll" (ByVal dc As Long, ByVal pos As Long) As Long
    Private Declare Function daveGetU8At Lib "libnodave.dll" (ByVal dc As Long, ByVal pos As Long) As Long
    Private Declare Function daveGetS16At Lib "libnodave.dll" (ByVal dc As Long, ByVal pos As Long) As Long
    Private Declare Function daveGetU16At Lib "libnodave.dll" (ByVal dc As Long, ByVal pos As Long) As Long
    Private Declare Function daveGetS32At Lib "libnodave.dll" (ByVal dc As Long, ByVal pos As Long) As Long
    '
    ' Is there an unsigned long? Or a longer integer than long? This doesn't work.
    'Private Declare Function daveGetU32At Lib "libnodave.dll" (ByVal dc As Long, ByVal pos As Long) As Long
    Private Declare Function daveGetFloatAt Lib "libnodave.dll" (ByVal dc As Long, ByVal pos As Long) As Single
    '
    ' Copy and convert a value of 8,16,or 32 bit, signed or unsigned into a buffer. The buffer
    ' is usually used by daveWriteBytes, daveWriteBits later.
    '
    Private Declare Function davePut8 Lib "libnodave.dll" (ByRef buffer As Byte, ByVal Value As Integer) As Integer
    Private Declare Function davePut16 Lib "libnodave.dll" (ByRef buffer As Byte, ByVal Value As Integer) As Integer
    Private Declare Function davePut32 Lib "libnodave.dll" (ByRef buffer As Byte, ByVal Value As Integer) As Integer
    Private Declare Function davePutFloat Lib "libnodave.dll" (ByRef buffer As Byte, ByVal Value As Single) As Integer
    '
    ' Copy and convert a value of 8,16,or 32 bit, signed or unsigned to position pos of a buffer.
    ' The buffer is usually used by daveWriteBytes, daveWriteBits later.
    '
    Private Declare Function davePut8At Lib "libnodave.dll" (ByRef buffer As Byte, ByVal pos As Long, ByVal Value As Long) As Long
    Private Declare Function davePut16At Lib "libnodave.dll" (ByRef buffer As Byte, ByVal pos As Long, ByVal Value As Long) As Long
    Private Declare Function davePut32At Lib "libnodave.dll" (ByRef buffer As Byte, ByVal pos As Long, ByVal Value As Long) As Long
    Private Declare Function davePutFloatAt Lib "libnodave.dll" (ByRef buffer As Byte, ByVal pos As Long, ByVal Value As Single) As Long
    '
    ' Takes a timer value and converts it into seconds:
    '
    Private Declare Function daveGetSeconds Lib "libnodave.dll" (ByVal dc As Long) As Single
    Private Declare Function daveGetSecondsAt Lib "libnodave.dll" (ByVal dc As Long, ByVal pos As Long) As Single
    '
    ' Takes a counter value and converts it to integer:
    '
    Private Declare Function daveGetCounterValue Lib "libnodave.dll" (ByVal dc As Long) As Long
    Private Declare Function daveGetCounterValueAt Lib "libnodave.dll" (ByVal dc As Long, ByVal pos As Long) As Long
    '
    ' Get the order code (MLFB number) from a PLC. Does NOT work with 200 family.
    '
    Private Declare Function daveGetOrderCode Lib "libnodave.dll" (ByVal en As Long, ByRef buffer As Byte) As Long
    '
    ' Connect to a PLC.
    '
    Private Declare Function daveConnectPLC Lib "libnodave.dll" (ByVal dc As Integer) As Integer
    '
    '
    ' Read a value or a block of values from PLC.
    '
    Private Declare Function daveReadBytes Lib "libnodave.dll" (ByVal dc As Integer, ByVal area As Integer, ByVal AreaNumber As Integer, ByVal start As Integer, ByVal numBytes As Integer, ByVal buffer As Integer) As Integer
    '
    ' Read a long block of values from PLC. Long means too long to transport in a single PDU.
    '
    Private Declare Function daveManyReadBytes Lib "libnodave.dll" (ByVal dc As Long, ByVal area As Long, ByVal AreaNumber As Long, ByVal start As Long, ByVal numBytes As Long, ByVal buffer As Long) As Long
    '
    ' Write a value or a block of values to PLC.
    '
    Private Declare Function daveWriteBytes Lib "libnodave.dll" (ByVal dc As Integer, ByVal area As Integer, ByVal AreaNumber As Integer, ByVal start As Integer, ByVal numBytes As Integer, ByRef buffer As Byte) As Integer
    '
    ' Write a long block of values to PLC. Long means too long to transport in a single PDU.
    '
    Private Declare Function daveWriteManyBytes Lib "libnodave.dll" (ByVal dc As Long, ByVal area As Long, ByVal AreaNumber As Long, ByVal start As Long, ByVal numBytes As Long, ByRef buffer As Byte) As Long
    '
    ' Read a bit from PLC. numBytes must be exactly one with all PLCs tested.
    ' Start is calculated as 8*byte number+bit number.
    '
    Private Declare Function daveReadBits Lib "libnodave.dll" (ByVal dc As Long, ByVal area As Long, ByVal AreaNumber As Long, ByVal start As Long, ByVal numBytes As Long, ByVal buffer As Long) As Long
    '
    ' Write a bit to PLC. numBytes must be exactly one with all PLCs tested.
    '
    Private Declare Function daveWriteBits Lib "libnodave.dll" (ByVal dc As Long, ByVal area As Long, ByVal AreaNumber As Long, ByVal start As Long, ByVal numBytes As Long, ByRef buffer As Byte) As Long
    '
    ' Set a bit in PLC to 1.
    '
    Private Declare Function daveSetBit Lib "libnodave.dll" (ByVal dc As Long, ByVal area As Long, ByVal AreaNumber As Long, ByVal start As Long, ByVal byteAddress As Long, ByVal bitAddress As Long) As Long
    '
    ' Set a bit in PLC to 0.
    '
    Private Declare Function daveClrBit Lib "libnodave.dll" (ByVal dc As Long, ByVal area As Long, ByVal AreaNumber As Long, ByVal start As Long, ByVal byteAddress As Long, ByVal bitAddress As Long) As Long
    '
    ' Read a diagnostic list (SZL) from PLC. Does NOT work with 200 family.
    '
    Private Declare Function daveReadSZL Lib "libnodave.dll" (ByVal dc As Long, ByVal ID As Long, ByVal index As Long, ByRef buffer As Byte) As Long
    '
    Private Declare Function daveListBlocksOfType Lib "libnodave.dll" (ByVal dc As Long, ByVal typ As Long, ByRef buffer As Byte) As Long
    Private Declare Function daveListBlocks Lib "libnodave.dll" (ByVal dc As Long, ByRef buffer As Byte) As Long
    Private Declare Function internalDaveGetBlockInfo Lib "libnodave.dll" Alias "daveGetBlockInfo" (ByVal dc As Long, ByRef buffer As Byte, ByVal block_type As Long, ByVal number As Long) As Long
    '
    Private Declare Function daveGetProgramBlock Lib "libnodave.dll" (ByVal dc As Long, ByVal blockType As Long, ByVal number As Long, ByRef buffer As Byte, ByRef Length As Long) As Long
    '
    ' Start or Stop a PLC:
    '
    Private Declare Function daveStart Lib "libnodave.dll" (ByVal dc As Long) As Long
    Private Declare Function daveStop Lib "libnodave.dll" (ByVal dc As Long) As Long
    '
    ' Set outputs (digital or analog ones) of an S7-200 that is in stop mode:
    '
    Private Declare Function daveForce200 Lib "libnodave.dll" (ByVal dc As Long, ByVal area As Long, ByVal start As Long, ByVal Value As Long) As Long
    '
    ' Initialize a multivariable read request.
    ' The parameter PDU must have been obtained from daveNew PDU:
    '
    Private Declare Sub davePrepareReadRequest Lib "libnodave.dll" (ByVal dc As Long, ByVal pdu As Long)
    '
    ' Add a new variable to a prepared request:
    '
    Private Declare Sub daveAddVarToReadRequest Lib "libnodave.dll" (ByVal pdu As Long, ByVal area As Long, ByVal AreaNumber As Long, ByVal start As Long, ByVal numBytes As Long)
    '
    ' Executes the entire request:
    '
    Private Declare Function daveExecReadRequest Lib "libnodave.dll" (ByVal dc As Long, ByVal pdu As Long, ByVal Rs As Long) As Long
    '
    ' Use the n-th result. This lets the functions daveGet<data type> work on that part of the
    ' internal buffer that contains the n-th result:
    '
    Private Declare Function daveUseResult Lib "libnodave.dll" (ByVal dc As Long, ByVal Rs As Long, ByVal resultNumber As Long) As Long
    '
    ' Frees the memory occupied by single results in the result structure. After that, you can reuse
    ' the resultSet in another call to daveExecReadRequest.
    '
    Private Declare Sub daveFreeResults Lib "libnodave.dll" (ByVal Rs As Long)
    '
    ' Adds a new bit variable to a prepared request. As with daveReadBits, numBytes must be one for
    ' all tested PLCs.
    '
    Private Declare Sub daveAddBitVarToReadRequest Lib "libnodave.dll" (ByVal pdu As Long, ByVal area As Long, ByVal AreaNumber As Long, ByVal start As Long, ByVal numBytes As Long)
    '
    ' Initialize a multivariable write request.
    ' The parameter PDU must have been obtained from daveNew PDU:
    '
    Private Declare Sub davePrepareWriteRequest Lib "libnodave.dll" (ByVal dc As Long, ByVal pdu As Long)
    '
    ' Add a new variable to a prepared write request:
    '
    Private Declare Sub daveAddVarToWriteRequest Lib "libnodave.dll" (ByVal pdu As Long, ByVal area As Long, ByVal AreaNumber As Long, ByVal start As Long, ByVal numBytes As Long, ByRef buffer As Byte)
    '
    ' Add a new bit variable to a prepared write request:
    '
    Private Declare Sub daveAddBitVarToWriteRequest Lib "libnodave.dll" (ByVal pdu As Long, ByVal area As Long, ByVal AreaNumber As Long, ByVal start As Long, ByVal numBytes As Long, ByRef buffer As Byte)
    '
    ' Execute the entire write request:
    '
    Private Declare Function daveExecWriteRequest Lib "libnodave.dll" (ByVal dc As Long, ByVal pdu As Long, ByVal Rs As Long) As Long
    '
    ' Initialize an MPI Adapter or NetLink Ethernet MPI gateway.
    ' While some protocols do not need this, I recommend to allways use it. It will do nothing if
    ' the protocol doesn't need it. But you can change protocols without changing your program code.
    '
    Private Declare Function daveInitAdapter Lib "libnodave.dll" (ByVal di As Integer) As Integer
    '
    ' Disconnect from a PLC. While some protocols do not need this, I recommend to allways use it.
    ' It will do nothing if the protocol doesn't need it. But you can change protocols without
    ' changing your program code.
    '
    Private Declare Function daveDisconnectPLC Lib "libnodave.dll" (ByVal dc As Long) As Long
    '
    '
    ' Disconnect from an MPI Adapter or NetLink Ethernet MPI gateway.
    ' While some protocols do not need this, I recommend to allways use it.
    ' It will do nothing if the protocol doesn't need it. But you can change protocols without
    ' changing your program code.
    '
    Private Declare Function daveDisconnectAdapter Lib "libnodave.dll" (ByVal dc As Long) As Long
    '
    '
    ' List nodes on an MPI or Profibus Network:
    '
    Private Declare Function daveListReachablePartners Lib "libnodave.dll" (ByVal dc As Long, ByRef buffer As Byte) As Long
    '
    '
    ' Set/change the timeout for an interface:
    '
    Private Declare Sub daveSetTimeout Lib "libnodave.dll" (ByVal di As Integer, ByVal maxTime As Integer)
    '
    ' Read the timeout setting for an interface:
    '
    Private Declare Function daveGetTimeout Lib "libnodave.dll" (ByVal di As Long)
    '
    ' Get the name of an interface. Do NOT use this, but the wrapper function defined below!
    '
    Private Declare Function daveInternalGetName Lib "libnodave.dll" Alias "daveGetName" (ByVal en As Long) As Long
    '
    ' Get the MPI address of a connection.
    '
    Private Declare Function daveGetMPIAdr Lib "libnodave.dll" (ByVal dc As Long) As Long
    '
    ' Get the length (in bytes) of the last data received on a connection.
    '
    Private Declare Function daveGetAnswLen Lib "libnodave.dll" (ByVal dc As Long) As Long
    '
    ' Get the maximum length of a communication packet (PDU).
    ' This value depends on your CPU and connection type. It is negociated in daveConnectPLC.
    ' A simple read can read MaxPDULen-18 bytes.
    '
    Private Declare Function daveGetMaxPDULen Lib "libnodave.dll" (ByVal dc As Long) As Long
    '
    ' Reserve memory for a resultSet and get a handle to it:
    '
    Private Declare Function daveNewResultSet Lib "libnodave.dll" () As Long
    '
    ' Destroy handles to daveInterface, daveConnections, PDUs and resultSets
    ' Free the memory reserved for them.
    '
    Private Declare Sub daveFree Lib "libnodave.dll" (ByVal Item As Long)
    '
    ' Reserve memory for a PDU and get a handle to it:
    '
    Private Declare Function daveNewPDU Lib "libnodave.dll" () As Long
    '
    ' Get the error code of the n-th single result in a result set:
    '
    Private Declare Function daveGetErrorOfResult Lib "libnodave.dll" (ByVal resultSet As Long, ByVal resultNumber As Long) As Long
    '
    Private Declare Function daveForceDisconnectIBH Lib "libnodave.dll" (ByVal di As Long, ByVal src As Long, ByVal dest As Long, ByVal mpi As Long) As Long
    '
    ' Helper functions to open serial ports and IP connections. You can use others if you want and
    ' pass their results to daveNewInterface.
    '
    ' Open a serial port using name, baud rate and parity. Everything else is set automatically:
    '
    Private Declare Function setPort Lib "libnodave.dll" (ByVal portName As String, ByVal BaudRate As String, ByVal parity As Byte) As Integer
    '
    ' Open a TCP/IP connection using port number (1099 for NetLink, 102 for ISO over TCP) and
    ' IP address. You must use an IP address, NOT a hostname!
    '
    Private Declare Function openSocket Lib "libnodave.dll" (ByVal port As Long, ByVal peer As String) As Long
    '
    ' Open an access oint. This is a name in you can add in the "set Programmer/PLC interface" dialog.
    ' To the access point, you can assign an interface like MPI adapter, CP511 etc.
    '
    Private Declare Function openS7online Lib "libnodave.dll" (ByVal peer As String) As Long
    '
    ' Close connections and serial ports opened with above functions:
    '
    Private Declare Function closePort Lib "libnodave.dll" (ByVal fh As Long) As Long
    '
    ' Close handle opende by opens7online:
    '
    Private Declare Function closeS7online Lib "libnodave.dll" (ByVal fh As Long) As Long
    '
    ' Read Clock time from PLC:
    '
    Private Declare Function daveReadPLCTime Lib "libnodave.dll" (ByVal dc As Long) As Long
    '
    ' set clock to a value given by user
    '
    Private Declare Function daveSetPLCTime Lib "libnodave.dll" (ByVal dc As Long, ByRef timestamp As Byte) As Long
    '
    ' set clock to PC system clock:
    '
    Private Declare Function daveSetPLCTimeToSystime Lib "libnodave.dll" (ByVal dc As Long) As Long
    '
    '       BCD conversions:
    '
    Private Declare Function daveToBCD Lib "libnodave.dll" (ByVal dc As Long) As Long
    Private Declare Function daveFromBCD Lib "libnodave.dll" (ByVal dc As Long) As Long
    '
    ' Here comes the wrapper code for functions returning strings:
    '
    Private Function daveStrError(ByVal code As Long) As String
        Dim sTmp As New String(Chr(0), 256)
        Dim ip As Long
        'sTmp = String$(256, 0)            'create a string of sufficient capacity
        ip = daveInternalStrerror(code)    ' have the text for code copied in
        Call daveStringCopy(ip, sTmp)    ' have the text for code copied in
        sTmp = Left$(sTmp, InStr(sTmp, Chr(0)) - 1) ' adjust the length
        daveStrError = sTmp                       ' and return result
    End Function

    Private Function daveAreaName(ByVal code As Long) As String
        'Dim sTmp As String
        Dim sTmp As New String(Chr(0), 256)
        Dim ip As Long
        'sTmp = String$(256, 0)            'create a string of sufficient capacity
        ip = daveInternalAreaName(code)    ' have the text for code copied in
        Call daveStringCopy(ip, sTmp)    ' have the text for code copied in
        sTmp = Left$(sTmp, InStr(sTmp, Chr(0)) - 1) ' adjust the length
        daveAreaName = sTmp                       ' and return result
    End Function
    Private Function daveBlockName(ByVal code As Long) As String
        ' Dim sTmp As String
        Dim sTmp As New String(Chr(0), 256)
        Dim ip As Long
        'sTmp = String$(256, 0)            'create a string of sufficient capacity
        ip = daveInternalBlockName(code)    ' have the text for code copied in
        Call daveStringCopy(ip, sTmp)    ' have the text for code copied in
        sTmp = Left$(sTmp, InStr(sTmp, Chr(0)) - 1) ' adjust the length
        daveBlockName = sTmp                       ' and return result
    End Function
    Private Function daveGetName(ByVal di As Long) As String
        'Dim sTmp As String
        Dim sTmp As New String(Chr(0), 256)
        Dim ip As Long
        ' sTmp = String$(256, 0)            'create a string of sufficient capacity
        ip = daveInternalGetName(di)    ' have the text for code copied in
        Call daveStringCopy(ip, sTmp)    ' have the text for code copied in
        sTmp = Left$(sTmp, InStr(sTmp, Chr(0)) - 1) ' adjust the length
        daveGetName = sTmp                       ' and return result
    End Function
    Private Function daveGetBlockInfo(ByVal di As Long) As Byte
        'Dim sTmp As String
        Dim sTmp As New String(Chr(0), 256)
        Dim ip As Long
        'sTmp = String$(256, 0)            'create a string of sufficient capacity
        ip = daveInternalGetName(di)    ' have the text for code copied in
        Call daveStringCopy(ip, sTmp)    ' have the text for code copied in
        sTmp = Left$(sTmp, InStr(sTmp, Chr(0)) - 1) ' adjust the length
        daveGetBlockInfo = sTmp                       ' and return result
    End Function
    '
    '*****************************************************
    ' End of interface declarations and helper functions.
    '*****************************************************
    '

    '
    ' Questa funzione server ad inizializzare il collegamento all hardware necessario
    ' per la connessione ai diversi PLC (1 o più)
    '
    Public Function initHardware(ByRef PLC As PLC_AdrType) As Integer
        Dim res As Long
        'Dim lStart As Long
        ' Inizializzo i valori
        'lStart = GetTickCount
        PLC.ph = 0
        PLC.dInterf = 0
        PLC.dConn = 0
        initHardware = -1
        res = -1

        ' Definisco parametri di comunicazione sia MPI che TCP
        PLC.sComParity = "O"

        If (PLC.iConnessione = daveProtoMPI) Then
            PLC.ph = setPort(PLC.sPortaSeriale, PLC.sBaudRate, Asc(Left$(PLC.sComParity, 1)))
        End If
        If (PLC.iConnessione = daveProtoMPI2) Then
            PLC.ph = setPort(PLC.sPortaSeriale, PLC.sBaudRate, Asc(Left$(PLC.sComParity, 1)))
        End If
        If (PLC.iConnessione = daveProtoMPI3) Then
            PLC.ph = setPort(PLC.sPortaSeriale, PLC.sBaudRate, Asc(Left$(PLC.sComParity, 1)))
        End If
        If (PLC.iConnessione = daveProtoPPI) Then
            PLC.sComParity = "E"
            PLC.ph = setPort(PLC.sPortaSeriale, PLC.sBaudRate, Asc(Left$(PLC.sComParity, 1)))
        End If
        If (PLC.iConnessione = daveProtoAS511) Then
            PLC.sComParity = "E"
            PLC.ph = setPort(PLC.sPortaSeriale, PLC.sBaudRate, Asc(Left$(PLC.sComParity, 1)))
        End If
        ' Apro socket per connessione ISO on TCP standard siemens
        If (PLC.iConnessione = daveProtoISOTCP) Then
            PLC.ph = openSocket(102, PLC.sPLC_IP)
        End If
        If (PLC.iConnessione = daveProtoISOTCP243) Then
            PLC.ph = openSocket(102, PLC.sPLC_IP)
        End If
        If (PLC.iConnessione = daveProtoMPI_IBH) Then
            PLC.ph = openSocket(1099, PLC.sPLC_IP)
        End If
        If (PLC.iConnessione = daveProtoPPI_IBH) Then
            PLC.ph = openSocket(1099, PLC.sPLC_IP)
        End If
        ' Apro collegamento S7OLINE to use Siemes libraries for transport (s7online)
        If (PLC.iConnessione = daveProtoS7online) Then
            PLC.ph = openS7online(PLC.sConnessione)
        End If

        ' si ritorna val > 0 porta OK
        If (PLC.ph > 0) Then
            PLC.dInterf = daveNewInterface(PLC.ph, PLC.ph, "IF1", 0, PLC.iConnessione, daveSpeed187k)

            'You can set longer timeout here, if you have  a slow connection
            If (PLC.iConnessione <= daveProtoAS511) Then
                Call daveSetTimeout(PLC.dInterf, 500000)
            End If
            res = daveInitAdapter(PLC.dInterf)
            If res = 0 Then
                initHardware = 0
            End If
        End If

        'PLC.lUsedTime = GetTickCount - lStart


    End Function

    '
    ' Questa funzione server per instaurare il collegamento con il PLC una volta
    ' che le connessioni hardware siano ok
    '
    Public Function initPlcConn(ByRef PLC As PLC_AdrType) As Integer
        Dim res As Long
        'Dim lStart As Long

        ' Inizializzo i valori
        'lStart = GetTickCount
        PLC.dConn = 0
        PLC.lCommErr = 0
        PLC.lCommOK = 0
        initPlcConn = -1
        res = -1

        ' si ritorna val > 0 porta OK
        If ((PLC.ph <> 0) And (PLC.dInterf <> 0)) Then
            '
            ' with ISO over TCP, set correct values for rack and slot of the CPU
            '
            PLC.dConn = daveNewConnection(PLC.dInterf, PLC.iMpiPpi, PLC.iRack, PLC.iSlot)
            res = daveConnectPLC(PLC.dConn)
            If res = 0 Then
                initPlcConn = 0
            End If
        End If
        'PLC.lUsedTime = GetTickCount - lStart
    End Function

    '
    ' Disconnect from PLC, disconnect from Adapter, close the serial interface or TCP/IP socket
    '
    ' Private Function cleanUp(ByRef ph As Long, ByRef di As Long, ByRef dc As Long)
    Public Function cleanUp(ByRef PLC As PLC_AdrType)
        Dim res As Long
        'Dim lStart As Long

        ' Inizializzo i valori
        'lStart = GetTickCount

        'Debug.Print "cleanUp Started"
        If PLC.dConn <> 0 Then
            res = daveDisconnectPLC(PLC.dConn)
            Call daveFree(PLC.dConn)
        End If
        If PLC.dInterf <> 0 Then
            res = daveDisconnectAdapter(PLC.dInterf)
        End If
        Sleep(500)
        If PLC.ph <> 0 Then
            If (PLC.iConnessione = daveProtoS7online) Then
                res = closeS7online(PLC.ph)
            Else
                res = closePort(PLC.ph)
            End If
        End If

        ' Libero la memoria
        Call daveFree(PLC.dInterf)
        PLC.dConn = 0
        PLC.dInterf = 0
        PLC.ph = 0

        'Debug.Print "cleanUp Finished"
        'PLC.lUsedTime = GetTickCount - lStart
    End Function


    '
    ' read some values from FD0,FD4,FD8,FD12 (MD0,MD4,MD8,MD12 in german notation)
    '  to read from data block 12, you would need to write:
    '  daveReadBytes(dc, daveDB, 12, 0, 16, 0)
    '
    Sub readFromPLC(ByRef PLC As PLC_AdrType)
        Dim sTmp As String
        Dim res As Long
        Dim res2 As Long

        sTmp = "Testing PLC read" & vbCrLf

        res2 = daveReadBytes(PLC.dConn, daveFlags, 0, 0, 16, 0)
        Debug.Print("result from readBytes:")
        If res2 = 0 Then
            sTmp = sTmp & "MD0(DINT): " & daveGetS32(PLC.dConn) & vbCrLf
            sTmp = sTmp & "MD4(DINT): " & daveGetS32(PLC.dConn) & vbCrLf
            sTmp = sTmp & "MD8(DINT): " & daveGetS32(PLC.dConn) & vbCrLf
            sTmp = sTmp & "MD12(REAL): " & daveGetFloat(PLC.dConn) & vbCrLf
        Else
            sTmp = sTmp & "error: " & daveStrError(res2) & vbCrLf
        End If
    End Sub

    Public Function StartPLC(ByRef PLC As PLC_AdrType)
        'Dim lStart As Long

        ' Gestione errori
        'On Error GoTo HandleError

        ' Inizializzo i valori
        'lStart = GetTickCount

        PLC.retval = daveStart(PLC.dConn)
        PLC.sMessaggio = daveStrError(PLC.retval)
        'PLC.lUsedTime = GetTickCount - lStart
        'Finally:

        'Exit Function
        'HandleError:
        'PLC.retval = -999
        'PLC.sMessaggio = "ERRORE SCONOSCIUTO"
        ' Resume Finally
    End Function

    Public Function stopPLC(ByRef PLC As PLC_AdrType)
        'Dim lStart As Long

        ' Gestione errori
        'On Error GoTo HandleError

        PLC.retval = daveStop(PLC.dConn)
        PLC.sMessaggio = daveStrError(PLC.retval)
        'PLC.lUsedTime = GetTickCount - lStart
        'Finally:
        'Exit Function
        'HandleError:
        'PLC.retval = -999
        'PLC.sMessaggio = "ERRORE SCONOSCIUTO"
        'Resume Finally
    End Function
    Sub readOrderCode(ByRef PLC As PLC_AdrType)
        Dim buffer(50) As Byte
        Dim i As Integer
        Dim sTmp As String
        Dim res As Long

        res = initHardware(PLC)
        If res = 0 Then
            res = initPlcConn(PLC)
            If res = 0 Then
                res = daveGetOrderCode(PLC.dConn, buffer(0))
                If res = 0 Then
                    For i = 0 To daveOrderCodeSize - 2 'last character is chr$(0), don't copy it
                        sTmp = sTmp + Chr(buffer(i))
                    Next i
                Else
                    sTmp = daveStrError(res)
                End If
            End If
        End If
        Call cleanUp(PLC)
    End Sub

    Public Function readDiagnostic(ByRef PLC As PLC_AdrType, ByVal ID As Long, ByVal SLZ_index As Long) As String
        ' The internal buffer is not big enough for all SZL lists.
        ' You must provide a buffer of sufficient size.
        Dim buffer(4096) As Byte
        Dim sTmp As String
        Dim res As Long
        Dim al As Long
        Dim index As Integer
        Dim ItemLen As Integer
        Dim ItemCount As Integer
        Dim bpos As Integer
        Dim sDiag As String
        Dim i As Integer, j As Integer
        Dim lStart As Long

        ' Gestione errori
        'On Error GoTo HandleError

        sDiag = ""

        ' Inizializzo i valori
        'lStart = GetTickCount

        res = daveReadSZL(PLC.dConn, ID, SLZ_index, buffer(0))
        If res = 0 Then
            al = daveGetAnswLen(PLC.dConn)
            If (al >= 4) Then
                ID = daveGetU16from(buffer(0))
                index = daveGetU16from(buffer(2))
                If (al >= 8) Then
                    ItemLen = daveGetU16from(buffer(4))
                    ItemCount = daveGetU16from(buffer(6))
                    bpos = 8    ' remember buffer position
                    For i = 0 To ItemCount - 1
                        '                    sDiag = ""
                        For j = 0 To ItemLen - 1
                            sDiag = sDiag + Hex$(buffer(bpos)) + ","
                            bpos = bpos + 1
                        Next j
                    Next i
                End If

            End If
        Else
            sTmp = daveStrError(res)
        End If
        'Finally:
        readDiagnostic = sDiag
        'Exit Function
        'HandleError:
        'PLC.retval = -999
        'PLC.sMessaggio = "ERRORE SCONOSCIUTO"
        'sDiag = "Err"

        'Resume Finally
    End Function
    '
    ' This is a test for passing back strings from Libnodave to VB(A):
    '
    Sub stringtest()
        Dim i As Integer
        Dim sTestoErrore As String
        Dim sTmp As String
        Dim sNomeBlocco As String

        For i = 0 To 255
            sTestoErrore = daveStrError(i)
            sTmp = daveAreaName(i)
            sNomeBlocco = daveBlockName(i)
        Next i
    End Sub

    Sub readMultipleItemsFromPLC(ByRef PLC As PLC_AdrType)
        Dim resultSet As Long
        Dim pdu As Long
        '
        ' Call daveSetDebug(&HFFFF)
        ' You may wonder what sense it might make to set debug level, as you cannot see
        ' messages when you opened excel from Widows GUI.
        ' You can invoke Excel from the console or from a batch file with:
        ' <myPathToExcel>\Excel.Exe <MyPathToXLS-File>VBATest.XLS >ExcelOut
        ' This will start Excel with VBATest.XLS and all debug messages (and a few from Excel itself)
        ' go into the file ExcelOut.
        '
        Dim res As Long
        Dim sTmp As String

        res = initHardware(PLC)
        If res = 0 Then
            res = initPlcConn(PLC)
            If res = 0 Then
                pdu = daveNewPDU
                Call davePrepareReadRequest(PLC.dConn, pdu)
                Call daveAddVarToReadRequest(pdu, daveFlags, 0, 0, 4)
                Call daveAddVarToReadRequest(pdu, daveFlags, 0, 8, 8)
                resultSet = daveNewResultSet
                res = daveExecReadRequest(PLC.dConn, pdu, resultSet)
                If res = 0 Then
                    res = daveUseResult(PLC.dConn, resultSet, 0)
                    Debug.Print(daveGetS32(PLC.dConn))
                    res = daveUseResult(PLC.dConn, resultSet, 0)
                    Debug.Print(daveGetS32(PLC.dConn))
                    Debug.Print(daveGetFloat(PLC.dConn))
                    daveFreeResults(resultSet)
                Else
                    sTmp = daveStrError(res)
                End If
                daveFree(resultSet)
                daveFree(pdu)
            End If
        End If
        Call cleanUp(PLC)
    End Sub

    Sub writeMultipleItemsToPLC(ByRef PLC As PLC_AdrType)
        Dim resultSet As Long
        Dim pdu As Long
        Dim res As Long
        Dim sTmp As String
        Dim buffer As Byte

        res = initHardware(PLC)
        If res = 0 Then
            res = initPlcConn(PLC)
            If res = 0 Then
                pdu = daveNewPDU
                res = daveGetMaxPDULen(PLC.dConn)
                Call davePrepareWriteRequest(PLC.dConn, pdu)
                Call daveAddVarToWriteRequest(pdu, daveFlags, 0, 0, 4, buffer)
                Call daveAddVarToWriteRequest(pdu, daveDB, 6, 8, 8, buffer)
                resultSet = daveNewResultSet
                res = daveExecWriteRequest(PLC.dConn, pdu, resultSet)
                If res = 0 Then
                    res = daveGetErrorOfResult(resultSet, 0)
                    res = daveGetErrorOfResult(resultSet, 1)
                    daveFreeResults(resultSet)
                Else
                    sTmp = daveStrError(res)
                End If
                daveFree(resultSet)
                daveFree(pdu)
            End If
        End If
        Call cleanUp(PLC)
    End Sub
    Public Function readProgramBlock(ByRef PLC As PLC_AdrType) As String
        Dim buffer(3000) As Byte, Length As Long
        Dim sTmp As String
        Dim res As Long
        Dim bpos As Integer
        Dim sDiag As String
        Dim i As Integer, j As Integer
        Dim sBlock As String

        res = daveGetProgramBlock(PLC.dConn, Asc("8"), 1, buffer(0), Length)
        bpos = 0
        sTmp = "Contents of OB1:"
        Debug.Print(sTmp)
        For i = 0 To 1 + Int(Length / 16)
            sDiag = ""
            For j = 0 To 15
                sDiag = sDiag + Hex$(buffer(bpos)) + ","
                bpos = bpos + 1
            Next j
            sTmp = sDiag
            sBlock = sBlock & sDiag & vbCrLf
            Debug.Print(sTmp)
        Next i

        readProgramBlock = sBlock
    End Function
    Public Function PlcBlockList(ByRef PLC As PLC_AdrType) As String
        Dim buffer(20000) As Byte
        Dim res As Long
        Dim i As Integer
        Dim sBlock As String

        sBlock = "List " & vbCrLf
        res = daveListBlocks(PLC.dConn, buffer(0))
        res = res
        For i = 0 To res
            sBlock = sBlock & "i=" & res & ",buffer("

        Next i
    End Function

    Public Function WriteMerker_PLC(ByRef PLC As PLC_AdrType, ByVal StartAdr As Integer, ByVal ElemType As Integer, ByVal WriteData As Object) As Long
        Dim buffer() As Byte
        Dim iNumItem As Integer
        Dim iNumBytes As Integer
        Dim i As Integer
        Dim ind As Integer
        Dim lStart As Long

        Dim abVettoreWrite() As Byte
        Dim iOffset As Integer
        Dim iLastByte As Integer
        Dim iByteToWrite As Integer

        ' Gestione errori
        'On Error GoTo HandleError

        ' Inizializzo i valori
        'lStart = GetTickCount

        ' Controllo tipologia dato
        If (ElemType < 1) Or (ElemType > 4) Then
            PLC.retval = -1001
        Else
            ' Controllo numero elemnti nel vettore
            iNumItem = UBound(WriteData) - LBound(WriteData) + 1
            iNumBytes = iNumItem * ElemType

            ReDim buffer(0 To iNumBytes + 10)
            ind = 0
            For i = 0 To iNumItem - 1
                Select Case ElemType
                    Case 1
                        Call davePut8(buffer(ind), WriteData(i))
                        ind = ind + ElemType
                    Case 2
                        Call davePut16(buffer(ind), WriteData(i))
                        ind = ind + ElemType
                    Case 4
                        Call davePut32(buffer(ind), WriteData(i))
                        ind = ind + ElemType
                End Select
            Next i
            ' Eseguo scritture ad un massimo di 220 byte alla volta
            iOffset = 0
            iLastByte = 0 + iNumBytes - 1
            iByteToWrite = br1MaxByteWrite
            Do
                If (iOffset + iByteToWrite > iLastByte) Then
                    iByteToWrite = iLastByte - iOffset + 1
                End If
                ReDim abVettoreWrite(0 To iByteToWrite - 1)
                Call CopyMemory(abVettoreWrite(0), buffer(iOffset), iByteToWrite)
                PLC.retval = daveWriteBytes(PLC.dConn, daveFlags, 0, StartAdr + iOffset, iByteToWrite, abVettoreWrite(0))
                PLC.sMessaggio = daveStrError(PLC.retval)
                iOffset = iOffset + iByteToWrite
            Loop While (iOffset < iLastByte) And (PLC.retval = 0)
        End If
        'PLC.lUsedTime = GetTickCount - lStart

        'Finally:
        WriteMerker_PLC = PLC.retval
        'Exit Function
        'HandleError:
        'PLC.retval = -999
        'PLC.sMessaggio = "ERRORE SCONOSCIUTO"
        'Resume Finally
    End Function

    Public Function WriteDB_PLC(ByRef PLC As PLC_AdrType, ByVal NumDB As Long, ByVal StartAdr As Integer, ByVal ElemType As Integer, ByVal WriteData As Object) As Long
        Dim buffer() As Byte
        Dim iNumItem As Integer
        Dim iNumBytes As Integer
        Dim i As Integer
        Dim ind As Integer
        Dim lStart As Long

        Dim abVettoreWrite() As Byte
        Dim iOffset As Integer
        Dim iLastByte As Integer
        Dim iByteToWrite As Integer

        ' Gestione errori
        '     On Error GoTo HandleError

        ' Inizializzo i valori
        '      lStart = GetTickCount

        ' Controllo tipologia dato
        If (ElemType < 1) Or (ElemType > 4) Then
            PLC.retval = -1001
        Else
            ' Controllo numero elemnti nel vettore
            iNumItem = UBound(WriteData) - LBound(WriteData) + 1
            iNumBytes = iNumItem * ElemType
            ReDim buffer(0 To iNumBytes + 10)
            ind = 0
            For i = 0 To iNumItem - 1
                Select Case ElemType
                    Case 1
                        Call davePut8(buffer(ind), WriteData(i))
                        ind = ind + ElemType
                    Case 2
                        Call davePut16(buffer(ind), WriteData(i))
                        ind = ind + ElemType
                    Case 4
                        Call davePut32(buffer(ind), WriteData(i))
                        ind = ind + ElemType
                End Select
            Next i
            ' Eseguo scritture ad un massimo di 220 byte alla volta
            iOffset = 0
            iLastByte = 0 + iNumBytes - 1
            iByteToWrite = br1MaxByteWrite
            Do
                If (iOffset + iByteToWrite > iLastByte) Then
                    iByteToWrite = iLastByte - iOffset + 1
                End If
                ReDim abVettoreWrite(0 To iByteToWrite - 1)
                Call CopyMemory(abVettoreWrite(0), buffer(iOffset), iByteToWrite)
                PLC.retval = daveWriteBytes(PLC.dConn, daveDB, NumDB, StartAdr + iOffset, iByteToWrite, abVettoreWrite(0))
                PLC.sMessaggio = daveStrError(PLC.retval)
                iOffset = iOffset + iByteToWrite
            Loop While (iOffset < iLastByte) And (PLC.retval = 0)

        End If
        '       PLC.lUsedTime = GetTickCount - lStart
        'Finally:
        WriteDB_PLC = PLC.retval
        '           Exit Function
        'HandleError:
        '      PLC.retval = -999
        '       PLC.sMessaggio = "ERRORE SCONOSCIUTO"
        'Resume Finally
    End Function


    Public Function ReadDB_PLC(ByRef PLC As PLC_AdrType, ByVal NumDB As Long, ByVal StartAdr As Integer, ByVal NumElem As Integer, ByVal ElemType As Integer, ByVal ReadData As Object) As Long
        Dim i As Integer
        Dim ind As Integer
        Dim lStart As Long
        Dim iNumBytes As Integer
        Dim iOffset As Integer
        Dim iLastByte As Integer
        Dim iByteToRead As Integer

        ' Gestione errori
        '      On Error GoTo HandleError

        ' Inizializzo i valori
        '     lStart = GetTickCount

        ' Controllo tipologia dato
        If (ElemType < 1) Or (ElemType > 4) Then
            PLC.retval = -1001
        Else
            ' Controllo numero elemnti nel vettore
            iNumBytes = NumElem * ElemType
            ' Eseguo letture ad un massimo di 220 byte alla volta
            iOffset = 0
            iLastByte = 0 + iNumBytes - 1
            iByteToRead = br1MaxByteRead
            i = 0
            ind = LBound(ReadData)
            Do
                If (iOffset + iByteToRead > iLastByte) Then
                    iByteToRead = iLastByte - iOffset + 1
                End If
                ' Leggo da PLC
                PLC.retval = daveReadBytes(PLC.dConn, daveDB, NumDB, StartAdr + iOffset, iByteToRead, 0)
                PLC.sMessaggio = daveStrError(PLC.retval)
                ' Se lettura OK sistemo vettore dato come argomento
                If (PLC.retval = 0) Then
                    For i = 0 To (iByteToRead / ElemType) - 1
                        Select Case ElemType
                            Case 1
                                ReadData(ind) = daveGetU8(PLC.dConn)
                            Case 2
                                ReadData(ind) = daveGetS16(PLC.dConn)
                            Case 4
                                ReadData(ind) = daveGetS32(PLC.dConn)
                        End Select
                        ind = ind + 1
                    Next i
                End If
                iOffset = iOffset + iByteToRead
            Loop While (iOffset < iLastByte) And (PLC.retval = 0)
        End If
        '       PLC.lUsedTime = GetTickCount - lStart

        'Finally:
        ReadDB_PLC = PLC.retval
        '           Exit Function
        'HandleError:
        '      PLC.retval = -999
        '       PLC.sMessaggio = "Errore generico"
        'Resume Finally
    End Function
    Public Function ReadMERKER_PLC(ByRef PLC As PLC_AdrType, ByVal StartAdr As Integer, ByVal NumElem As Integer, ByVal ElemType As Integer, ByVal ReadData As Object) As Long
        Dim i As Integer
        Dim ind As Integer
        Dim lStart As Long
        Dim iNumBytes As Integer
        Dim iOffset As Integer
        Dim iLastByte As Integer
        Dim iByteToRead As Integer


        ' Gestione errori
        '     On Error GoTo HandleError

        ' Inizializzo i valori
        '      lStart = GetTickCount

        ' Controllo tipologia dato
        If (ElemType < 1) Or (ElemType > 4) Then
            PLC.retval = -1001
        Else
            ' Controllo numero elemnti nel vettore
            iNumBytes = NumElem * ElemType
            ' Eseguo letture ad un massimo di 220 byte alla volta
            iOffset = 0
            iLastByte = 0 + iNumBytes - 1
            iByteToRead = br1MaxByteRead
            i = 0
            ind = LBound(ReadData)
            Do
                If (iOffset + iByteToRead > iLastByte) Then
                    iByteToRead = iLastByte - iOffset + 1
                End If
                ' Leggo da PLC
                PLC.retval = daveReadBytes(PLC.dConn, daveFlags, 0, StartAdr + iOffset, iByteToRead, 0)
                PLC.sMessaggio = daveStrError(PLC.retval)
                ' Se lettura OK sistemo vettore dato come argomento
                If (PLC.retval = 0) Then
                    For i = 0 To (iByteToRead / ElemType) - 1
                        Select Case ElemType
                            Case 1
                                ReadData(ind) = daveGetU8(PLC.dConn)
                            Case 2
                                ReadData(ind) = daveGetS16(PLC.dConn)
                            Case 4
                                ReadData(ind) = daveGetS32(PLC.dConn)
                        End Select
                        ind = ind + 1
                    Next i
                End If
                iOffset = iOffset + iByteToRead
            Loop While (iOffset < iLastByte) And (PLC.retval = 0)
        End If
        '       PLC.lUsedTime = GetTickCount - lStart

        'Finally:
        ReadMERKER_PLC = PLC.retval
        '           Exit Function
        'HandleError:
        '      PLC.retval = -999
        '       PLC.sMessaggio = "Errore generico"
        'Resume Finally
    End Function

    Public Function InitConnPlcSiemens(ByRef PLC As PLC_AdrType)

        PLC.sPortaSeriale = "COM1"
        PLC.sBaudRate = "9600"
        PLC.iMpiPpi = 2

        'Sistemo i Dati Fissi
        PLC.iRack = 0
        PLC.iSlot = 2

        'Tipologia di Connessione / Default = 10 (PPI For S7 200)
        PLC.iConnessione = 10
        PLC.sConnessione = "/S7ONLINE"

    End Function

    Public Function ApriConnPlcSiemens(ByRef PLC As PLC_AdrType) As Boolean
        'Funzione di Apertura della DLL
        'Restituisco eventualmente TRUE se ho aperto bene, FALSE se ho avuto problemi

        'Verifico che non vi era gia una Apertura in corso
        If (PLC.dConn > 0) Then

            ' Richiamo Disconnessione
            Call modSiemens.cleanUp(PLC)
            ApriConnPlcSiemens = False
        Else
            ' Richiamo connessione
            If (modSiemens.initHardware(PLC) < 0) Then
                ApriConnPlcSiemens = False
            Else
                Call modSiemens.initPlcConn(PLC)
                ApriConnPlcSiemens = True
            End If
        End If

    End Function
    Public Function StopPLCSiemens(ByVal PLC As PLC_AdrType) As Boolean

        'Funzione per mettere in STOP la CPU
        Call modSiemens.stopPLC(PLC)

        'Verifico se mi e' tornato indietro un errore
        If PLC.retval = 0 Then
            StopPLCSiemens = True
        Else
            StopPLCSiemens = False
        End If
    End Function
    Public Function StartPLCSiemens(ByVal PLC As PLC_AdrType) As Boolean

        'Funzione per mettere in START la CPU
        Call modSiemens.StartPLC(PLC)

        'Verifico se mi e' tornato indietro un errore
        If PLC.retval = 0 Then
            StartPLCSiemens = True
        Else
            StartPLCSiemens = False
        End If
    End Function
    Public Function LeggiDB(ByVal numeroDB As Long, ByVal startDB As Integer, ByVal numeroElementi As Integer, ByRef Risultato() As Byte, ByVal PLC As PLC_AdrType) As Boolean
        'Funzione per leggere i DB dal PLC
        'Vengono chiamati DB ma su l'S7200 in realtà sono i VB.

        'Ridimensiono il Vettore per il numero di campi desiderati
        ReDim Risultato(0 To numeroElementi - 1)

        'Leggo dal PLC il tutto
        Call ReadDB_PLC(PLC, numeroDB, startDB, numeroElementi, 1, Risultato)

        'Verifico se e' andata a buon fine la lettura
        If (PLC.retval = 0) Then
            LeggiDB = True
        Else
            LeggiDB = False
        End If

    End Function

    Public Function ScriviDB(ByVal numeroDB As Long, ByVal startDB As Integer, ByVal ValoreDB As Integer, ByVal PLC As PLC_AdrType) As Boolean
        'Funzione per Scrivere un Valore sul DB del PLC
        Dim VettorePLC() As Byte
        ReDim VettorePLC(0 To 0)

        VettorePLC(0) = ValoreDB

        'Eseguo il comando di Write sul PLC
        Call WriteDB_PLC(PLC, numeroDB, startDB, 1, VettorePLC)

        'Verifico la risposta PLC
        If (PLC.retval = 0) Then
            ScriviDB = True
        Else
            ScriviDB = False
        End If
    End Function

    Public Function ScriviDBMultiplo(ByVal numeroDB As Long, ByVal startDB As Integer, ByVal VettoreDati() As Byte, ByVal PLC As PLC_AdrType) As Boolean

        'Eseguo il comando di Write sul PLC
        Call WriteDB_PLC(PLC, numeroDB, startDB, 1, VettoreDati)

        'Verifico la risposta PLC
        If (PLC.retval = 0) Then
            ScriviDBMultiplo = True
        Else
            ScriviDBMultiplo = False
        End If

    End Function
End Module
