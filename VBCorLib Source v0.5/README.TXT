Welcome to VBCorLib Release v0.5!

A new website has been constructed to help developers that are interested in
the VBCorLib library. It will continue to grow as new things emerge.
The site is www.kellyethridge.com/vbcorlib

What's new in v0.5?
The primary objective in this release was to implement the System.IO namespace.
There were some fixes and enhancements made to the existing code base as well.
- BinaryReader
- BinaryWriter
- Directory
- DirectoryInfo
- DirectoryNotFoundException
- EndOfStreamException
- File
- FileInfo
- FileStream
- FileSystemInfo
- MemoryStream (updated)
- Path (updated)
- Stream
- StreamReader
- StreamWriter
- TextReader
- TextWriter
- More documentation
- Fixed all of the ToArray functions.

What's new in v0.3?
- UTF8Encoding
- UTF7Encoding
- UnicodeEncoding (supports big-endian byte ordering)
- ASCIIEncoding
- Stream
- FileStream (supports files larger than 2gig bytes)
- MemoryStream
- StringReader
- StringWriter
- TextReader
- TextWriter
- Minor Fixes
- Updated/Fixed VBCultureBuilder
- More documenting (can't seem to catch up :)

Be sure to check out the Constructors class to see what you can create. All of 
the function in Constructors are global-multiuse accessable.

What it is:
VBCorLib is an ongoing project to recreate the .NET mscorlib assembly for VB6.


Many of the utility classes in mscorlib can be recreated nearly identically in
VB6. Those portions that can't, are simulated as best as can be. 

The building of the massive library carries a set of Unit Tests that can be
selectively performed to ensure the integrity of the project. In order to run
any of the tests, SimplyVBUnit v1.4.x must be installed on the system. This
can be obtained at www.sourceforge.net/projects/simplyvbunit.

This project includes both interfaces and classes. Sometimes methods need to be
accessed as a static method in a class. Since VB6 does not support this, it is
simulated by creating a class to contain static methods and providing a public
function of the class name. This gives the equivalence of accessing static methods
in .NET classes. An example would be for using the Buffer static methods.

	i = Buffer.ByteLength(myArray)
	
Buffer is not an instantiated object. You simply access the static methods through
the class name, without any creation of an object on your part.

Not all of the classes in this project are a public. Many are helper classes for
such things as enumerating collection classes, or being a wrapper class to maintain
limited access to an underlying object. These classes are not listed here, but
should be considered as part of the parent class, so the status indicates the
completion of those helper classes as well.

Tools:
Now included with VBCorLib source is a new directory called Tools. In this
directory is a set of tools that helps build data external of VBCorLib. The
only tool currently present is VBCultureInfoBuilder.

VBCultureInfoBuilder:
In order for VBCorLib to support such a vast array of cultures is to maintain
culture specific information outside of the source code. A file named VBCultures.nlp
is mapped into memory to quickly load culture specific information when a new
CultureInfo object is created. The VBCultures.nlp file is built using
VBCultureInfoBuilder. The data that is used by the builder is in a subdirectory named
CultureInfoData. In it is a set of XML files containing the culture specific information.
Once VBCultures.nlp is built, it can be placed in the App.Path of the executing program
if you wish to have a special version, or it can be placed in the Windows directory. If
the file is not found, then only the Invariant culture type is supported.

BREAKING CHANGES SINCE 0.1 (Compatible with v0.2)
cArray.CreateInstance		Now creates an empty array of up to 3 dimensions.
				Added cArray.NewArray to create a filled 1 dimension array.

--- Unit Tests ---
An extensive set of tests is included in the release. In order to load and perform the
tests, SimplyVBUnit v1.4.x must be installed on the system. It can be obtained from
www.sourceforge.net/projects/simplyvbunit.

The tests not only perform and integrity check, but also show how to use the classes.

--- Type Library ---
The type library of VBCorLibTypes.tlb and VBVM6Lib.tlb must be registered using REGTLIB.EXE in order to load up the
project in VB. It is not required to use the compiled DLL. It is located in the 'VBCorLib\Type Libraries' folder.

Below is a list of interfaces and classes that are in the project. 
There are 3 columns:
	Name is the name of the interface or class.
	Status is the completion status
	Release is the version in which the interface or class was added to the project.

--- Interfaces --
Name						Status	Release
AsyncCallback					100%	0.3
Decoder						100%	0.3
Encoder						100%	0.3
Encoding					100%	0.3
ICloneable 					100%	0.1
ICollection 					100%	0.1
IComparable 					100%	0.1
IComparer					100%	0.1
ICustomFormatter				100%	0.1
IDictionary					100%	0.1
IDictionaryEnumerator				100%	0.1
IEnumerable					100%	0.1
IEnumerator					100%	0.1
IFormatProvider					100%	0.1
IFormattable					100%	0.1
IHashcodeProvider				100%	0.1
IList						100%	0.1
Stream						100%	0.3
TextReader					100%	0.3
TextWriter					100%	0.3

--- Instance Classes ---
Name						Status	Release
ArgumentException				100%	0.1
ArgumentNullException				100%	0.1
ArgumentOutOfRangeException			100%	0.1
ArrayList					100%	0.1
ArrayTypeMismatchException			100%	0.1
ASCIIEncoding					100%	0.3
BitArray					100%	0.1
cDateTime					100%	0.1
cObject						100%	0.1
Comparer					100%	0.1
CaseInsensitiveComparer				100%	0.1
CultureInfo					90%	0.2
DateTimeFormatInfo				100%	0.1
Exception					100%	0.1
FileStream					100%	0.3
FormatException					100%	0.1
Hashtable					100%	0.1
IndexOutOfRangeException			100%	0.1
InvalidCastException				100%	0.1
InvalidOperationException			100%	0.1
MappedFile					100%	0.2
MemoryStream					100%	0.3
NotSupportedException				100%	0.1
NumberFormatInfo				100%	0.1
OutOfMemoryException				100%	0.1
OperatingSystem					100%	0.2
PathTooLongException				100%	0.3
PlatformNotSupportedException			100%	0.2
Queue						100%	0.1
Random						100%	0.1
RankException					100%	0.1
SerializationException				100%	0.1
SortedList					100%	0.1
Stack						100%	0.1
StringBuilder					100%	0.1
SystemException					100%	0.1
TimeSpan					100%	0.1
UnicodeEncoding					100%	0.3
UTF7Encoding					100%	0.3
UTF8Encoding					100%	0.3
Version						100%	0.1
WeakReference					100%	0.1

--- Static Classes ---
These are used to access static methods by typing these class names
followed by a period, then the method name.
Examples would be:
	cArray.IsNull(myArray)
	cString.Format("{0}", "hello")

Name						Status	Release
ArrayList					100%	0.1
BitArray					100%	0.1
BitConverter					100%	0.1
Buffer						100%	0.1
cArray						100%	0.1
cDateTime					100%	0.1
CultureInfo					90%	0.2
Comparer					100%	0.1
cString						100%	0.1
DateTimeFormatInfo				100%	0.1
Encoding					95%	0.3
Environment					100%	0.1
NumberFormatInfo				100%	0.1
Path						100%	0.3
TextReader					100%	0.3
TimeSpan					100%	0.1
TimeZone					100%	0.2
Version						100%	0.1

--- Constructors ---
This is a set of global-multiuse functions that can be used to
help easily create objects with parameters. To see these functions
use the Cor function. An example would be:
	Set obj = Cor.NewArgumentException("Something is wrong.")
		or
	Set obj = NewArgumentException("Something is wrong.")

--- Exceptions ---
Exceptions can be thrown and caught using the Throw and Catch functions.

Throw is a function that accepts an object that implements the Exception
interface. By calling Throw with an exception object, the object is set
to be caught, then an error is raised using the information in the exception object.

Catch is a function that takes an exception variable and an optional Err object. It
returns a boolean indicating if an exception was caught. A simple use would be:
	Dim ex As Exception
	If Catch(ex) then
		MsgBox ex.ToString
	End If

--- Compiling ---
Compiling is performed in the normal manner. There is a constant defined in
'Conditional Compilation Arguments' that allows a Type to be present at the
top of the cArray class. Setting the INDEBUG to 0 will remove the Type, but also
prevent the associated test from passing.

--- To Use ---
Simply include a reference to the resulting DLL and everything will be available.
There is no initialization required, all static methods are fully available.

--- Documentation ---
Some documentation is written within the classes. They are in a format that is
understandable by VB.DOX document generator. More documentation will be added
as the project continues. If you wish to generate the documentation from the
classes using VB.DOX, you will need to download it from sourceforge.net.

--- Template Classes ---
These classes are meant as starting points for creating new classes to work within
the VBCorLib system. By implementing atleast cObject interface, your classes can
be used extensively within the system. 

cObjectBase template class provides a skeleton from which to build your class.
The default behaviours can be overriden to your specifications.

ExceptionBase template class is a starting point for creating exception classes that
are custom to your needs and can still be used by the system.

SystemExceptionBase template class is are more precise exception primarily used by the
set of system specific exception classes.

IOExceptionBase template class is the start in creating I/O error exceptions such
as the FileNotFoundException.