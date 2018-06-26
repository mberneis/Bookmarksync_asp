<%

'	AspUpload Constants Include File

'	Copyright (c) Persits Software, Inc. All rights reserved.



' LogonUser's Type Parameter
const		LOGON_INTERACTIVE	= 2
const		LOGON_NETWORK		= 3
const		LOGON_BATCH			= 4
const		LOGON_SERVICE		= 5


' Generic Access Types
const		GENERIC_ALL			= &H10000000
const		GENERIC_EXECUTE		= &H20000000
const		GENERIC_WRITE		= &H40000000
const		GENERIC_READ		= &H80000000

' Standard Access Types
const		DELETE				= &H00010000
const		READ_CONTROL		= &H00020000
const		WRITE_DAC			= &H00040000
const		WRITE_OWNER			= &H00080000
const		WRITE_SYNCHRONIZE	= &H00100000


' Specific Access Types for Files

const		FILE_GENERIC_READ		= &H120089
const		FILE_GENERIC_WRITE		= &H120116
const		FILE_GENERIC_EXECUTE	= &H1200A0

const		FILE_READ_DATA			= &H0001
const		FILE_WRITE_DATA			= &H0002
const		FILE_APPEND_DATA		= &H0004
const		FILE_READ_EA			= &H0008
const		FILE_WRITE_EA			= &H0010
const		FILE_EXECUTE			= &H0020
const		FILE_READ_ATTRIBUTES	= &H0080
const		FILE_WRITE_ATTRIBUTES	= &H0100


' File Attributes
const		FILE_ATTRIBUTE_READONLY		= &H1
const		FILE_ATTRIBUTE_HIDDEN		= &H2
const		FILE_ATTRIBUTE_SYSTEM		= &H4
const		FILE_ATTRIBUTE_DIRECTORY	= &H10
const		FILE_ATTRIBUTE_ARCHIVE		= &H20
const		FILE_ATTRIBUTE_NORMAL		= &H80
const		FILE_ATTRIBUTE_TEMPORARY	= &H100
const		FILE_ATTRIBUTE_COMPRESSED	= &H800

' Sort-by Attributes for Directory Collection.
' These are NOT standard Windows NT constants
const		SORTBY_NAME				=	1
const		SORTBY_TYPE				=	2
const		SORTBY_SIZE				=	3
const		SORTBY_CREATIONTIME		=	4
const		SORTBY_LASTWRITETIME	=	5
const		SORTBY_LASTACCESSTIME	=	6

%>
