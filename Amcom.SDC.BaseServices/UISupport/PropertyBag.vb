'******************************************************************* 
' * 
' * PropertyBag.cs 
' * -------------- 
' * Copyright (C) 2002 Tony Allowatt 
' * Last Update: 12/14/2002 
' * 
' * THE SOFTWARE IS PROVIDED BY THE AUTHOR "AS IS", WITHOUT WARRANTY 
' * OF ANY KIND, EXPRESS OR IMPLIED. IN NO EVENT SHALL THE AUTHOR BE 
' * LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY ARISING FROM, 
' * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OF THIS 
' * SOFTWARE. 
' * 
' * Public types defined in this file: 
' * ---------------------------------- 
' * namespace Flobbster.Windows.Forms 
' * class PropertySpec 
' * class PropertySpecEventArgs 
' * delegate PropertySpecEventHandler 
' * class PropertyBag 
' * class PropertyBag.PropertySpecCollection 
' * class PropertyTable 
' * 
' ******************************************************************* 


Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Drawing.Design

Namespace UISupport
    ''' <summary> 
    ''' Represents a single property in a PropertySpec. 
    ''' </summary> 
    Public Class PropertySpec
        Private m_attributes As Attribute()
        Private m_category As String
        Private m_defaultValue As Object
        Private m_description As String
        Private editor As String
        Private m_name As String
        Private type As String
        Private typeConverter As String

        ''' <summary> 
        ''' Initializes a new instance of the PropertySpec class. 
        ''' </summary> 
        ''' <param name="name">The name of the property displayed in the property grid.</param> 
        ''' <param name="type">The fully qualified name of the type of the property.</param> 
        Public Sub New(ByVal name As String, ByVal type As String)
            Me.New(name, type, Nothing, Nothing, Nothing)
        End Sub

        ''' <summary> 
        ''' Initializes a new instance of the PropertySpec class. 
        ''' </summary> 
        ''' <param name="name">The name of the property displayed in the property grid.</param> 
        ''' <param name="type">A Type that represents the type of the property.</param> 
        Public Sub New(ByVal name As String, ByVal type As Type)
            Me.New(name, type.AssemblyQualifiedName, Nothing, Nothing, Nothing)
        End Sub

        ''' <summary> 
        ''' Initializes a new instance of the PropertySpec class. 
        ''' </summary> 
        ''' <param name="name">The name of the property displayed in the property grid.</param> 
        ''' <param name="type">The fully qualified name of the type of the property.</param> 
        ''' <param name="category">The category under which the property is displayed in the 
        ''' property grid.</param> 
        Public Sub New(ByVal name As String, ByVal type As String, ByVal category As String)
            Me.New(name, type, category, Nothing, Nothing)
        End Sub

        ''' <summary> 
        ''' Initializes a new instance of the PropertySpec class. 
        ''' </summary> 
        ''' <param name="name">The name of the property displayed in the property grid.</param> 
        ''' <param name="type">A Type that represents the type of the property.</param> 
        ''' <param name="category"></param> 
        Public Sub New(ByVal name As String, ByVal type As Type, ByVal category As String)
            Me.New(name, type.AssemblyQualifiedName, category, Nothing, Nothing)
        End Sub

        ''' <summary> 
        ''' Initializes a new instance of the PropertySpec class. 
        ''' </summary> 
        ''' <param name="name">The name of the property displayed in the property grid.</param> 
        ''' <param name="type">The fully qualified name of the type of the property.</param> 
        ''' <param name="category">The category under which the property is displayed in the 
        ''' property grid.</param> 
        ''' <param name="description">A string that is displayed in the help area of the 
        ''' property grid.</param> 
        Public Sub New(ByVal name As String, ByVal type As String, ByVal category As String, ByVal description As String)
            Me.New(name, type, category, description, Nothing)
        End Sub

        ''' <summary> 
        ''' Initializes a new instance of the PropertySpec class. 
        ''' </summary> 
        ''' <param name="name">The name of the property displayed in the property grid.</param> 
        ''' <param name="type">A Type that represents the type of the property.</param> 
        ''' <param name="category">The category under which the property is displayed in the 
        ''' property grid.</param> 
        ''' <param name="description">A string that is displayed in the help area of the 
        ''' property grid.</param> 
        Public Sub New(ByVal name As String, ByVal type As Type, ByVal category As String, ByVal description As String)
            Me.New(name, type.AssemblyQualifiedName, category, description, Nothing)
        End Sub

        ''' <summary> 
        ''' Initializes a new instance of the PropertySpec class. 
        ''' </summary> 
        ''' <param name="name">The name of the property displayed in the property grid.</param> 
        ''' <param name="type">The fully qualified name of the type of the property.</param> 
        ''' <param name="category">The category under which the property is displayed in the 
        ''' property grid.</param> 
        ''' <param name="description">A string that is displayed in the help area of the 
        ''' property grid.</param> 
        ''' <param name="defaultValue">The default value of the property, or null if there is 
        ''' no default value.</param> 
        Public Sub New(ByVal name As String, ByVal type As String, ByVal category As String, ByVal description As String, ByVal defaultValue As Object)
            Me.m_name = name
            Me.type = type
            Me.m_category = category
            Me.m_description = description
            Me.m_defaultValue = defaultValue
            Me.m_attributes = Nothing
        End Sub

        ''' <summary> 
        ''' Initializes a new instance of the PropertySpec class. 
        ''' </summary> 
        ''' <param name="name">The name of the property displayed in the property grid.</param> 
        ''' <param name="type">A Type that represents the type of the property.</param> 
        ''' <param name="category">The category under which the property is displayed in the 
        ''' property grid.</param> 
        ''' <param name="description">A string that is displayed in the help area of the 
        ''' property grid.</param> 
        ''' <param name="defaultValue">The default value of the property, or null if there is 
        ''' no default value.</param> 
        Public Sub New(ByVal name As String, ByVal type As Type, ByVal category As String, ByVal description As String, ByVal defaultValue As Object)
            Me.New(name, type.AssemblyQualifiedName, category, description, defaultValue)
        End Sub

        ''' <summary> 
        ''' Initializes a new instance of the PropertySpec class. 
        ''' </summary> 
        ''' <param name="name">The name of the property displayed in the property grid.</param> 
        ''' <param name="type">The fully qualified name of the type of the property.</param> 
        ''' <param name="category">The category under which the property is displayed in the 
        ''' property grid.</param> 
        ''' <param name="description">A string that is displayed in the help area of the 
        ''' property grid.</param> 
        ''' <param name="defaultValue">The default value of the property, or null if there is 
        ''' no default value.</param> 
        ''' <param name="editor">The fully qualified name of the type of the editor for this 
        ''' property. This type must derive from UITypeEditor.</param> 
        ''' <param name="typeConverter">The fully qualified name of the type of the type 
        ''' converter for this property. This type must derive from TypeConverter.</param> 
        Public Sub New(ByVal name As String, ByVal type As String, ByVal category As String, ByVal description As String, ByVal defaultValue As Object, ByVal editor As String, _
        ByVal typeConverter As String)
            Me.New(name, type, category, description, defaultValue)
            Me.editor = editor
            Me.typeConverter = typeConverter
        End Sub

        ''' <summary> 
        ''' Initializes a new instance of the PropertySpec class. 
        ''' </summary> 
        ''' <param name="name">The name of the property displayed in the property grid.</param> 
        ''' <param name="type">A Type that represents the type of the property.</param> 
        ''' <param name="category">The category under which the property is displayed in the 
        ''' property grid.</param> 
        ''' <param name="description">A string that is displayed in the help area of the 
        ''' property grid.</param> 
        ''' <param name="defaultValue">The default value of the property, or null if there is 
        ''' no default value.</param> 
        ''' <param name="editor">The fully qualified name of the type of the editor for this 
        ''' property. This type must derive from UITypeEditor.</param> 
        ''' <param name="typeConverter">The fully qualified name of the type of the type 
        ''' converter for this property. This type must derive from TypeConverter.</param> 
        Public Sub New(ByVal name As String, ByVal type As Type, ByVal category As String, ByVal description As String, ByVal defaultValue As Object, ByVal editor As String, _
        ByVal typeConverter As String)
            Me.New(name, type.AssemblyQualifiedName, category, description, defaultValue, editor, _
            typeConverter)
        End Sub

        ''' <summary> 
        ''' Initializes a new instance of the PropertySpec class. 
        ''' </summary> 
        ''' <param name="name">The name of the property displayed in the property grid.</param> 
        ''' <param name="type">The fully qualified name of the type of the property.</param> 
        ''' <param name="category">The category under which the property is displayed in the 
        ''' property grid.</param> 
        ''' <param name="description">A string that is displayed in the help area of the 
        ''' property grid.</param> 
        ''' <param name="defaultValue">The default value of the property, or null if there is 
        ''' no default value.</param> 
        ''' <param name="editor">The Type that represents the type of the editor for this 
        ''' property. This type must derive from UITypeEditor.</param> 
        ''' <param name="typeConverter">The fully qualified name of the type of the type 
        ''' converter for this property. This type must derive from TypeConverter.</param> 
        Public Sub New(ByVal name As String, ByVal type As String, ByVal category As String, ByVal description As String, ByVal defaultValue As Object, ByVal editor As Type, _
        ByVal typeConverter As String)
            Me.New(name, type, category, description, defaultValue, editor.AssemblyQualifiedName, _
            typeConverter)
        End Sub

        ''' <summary> 
        ''' Initializes a new instance of the PropertySpec class. 
        ''' </summary> 
        ''' <param name="name">The name of the property displayed in the property grid.</param> 
        ''' <param name="type">A Type that represents the type of the property.</param> 
        ''' <param name="category">The category under which the property is displayed in the 
        ''' property grid.</param> 
        ''' <param name="description">A string that is displayed in the help area of the 
        ''' property grid.</param> 
        ''' <param name="defaultValue">The default value of the property, or null if there is 
        ''' no default value.</param> 
        ''' <param name="editor">The Type that represents the type of the editor for this 
        ''' property. This type must derive from UITypeEditor.</param> 
        ''' <param name="typeConverter">The fully qualified name of the type of the type 
        ''' converter for this property. This type must derive from TypeConverter.</param> 
        Public Sub New(ByVal name As String, ByVal type As Type, ByVal category As String, ByVal description As String, ByVal defaultValue As Object, ByVal editor As Type, _
        ByVal typeConverter As String)
            Me.New(name, type.AssemblyQualifiedName, category, description, defaultValue, editor.AssemblyQualifiedName, _
            typeConverter)
        End Sub

        ''' <summary> 
        ''' Initializes a new instance of the PropertySpec class. 
        ''' </summary> 
        ''' <param name="name">The name of the property displayed in the property grid.</param> 
        ''' <param name="type">The fully qualified name of the type of the property.</param> 
        ''' <param name="category">The category under which the property is displayed in the 
        ''' property grid.</param> 
        ''' <param name="description">A string that is displayed in the help area of the 
        ''' property grid.</param> 
        ''' <param name="defaultValue">The default value of the property, or null if there is 
        ''' no default value.</param> 
        ''' <param name="editor">The fully qualified name of the type of the editor for this 
        ''' property. This type must derive from UITypeEditor.</param> 
        ''' <param name="typeConverter">The Type that represents the type of the type 
        ''' converter for this property. This type must derive from TypeConverter.</param> 
        Public Sub New(ByVal name As String, ByVal type As String, ByVal category As String, ByVal description As String, ByVal defaultValue As Object, ByVal editor As String, _
        ByVal typeConverter As Type)
            Me.New(name, type, category, description, defaultValue, editor, _
            typeConverter.AssemblyQualifiedName)
        End Sub

        ''' <summary> 
        ''' Initializes a new instance of the PropertySpec class. 
        ''' </summary> 
        ''' <param name="name">The name of the property displayed in the property grid.</param> 
        ''' <param name="type">A Type that represents the type of the property.</param> 
        ''' <param name="category">The category under which the property is displayed in the 
        ''' property grid.</param> 
        ''' <param name="description">A string that is displayed in the help area of the 
        ''' property grid.</param> 
        ''' <param name="defaultValue">The default value of the property, or null if there is 
        ''' no default value.</param> 
        ''' <param name="editor">The fully qualified name of the type of the editor for this 
        ''' property. This type must derive from UITypeEditor.</param> 
        ''' <param name="typeConverter">The Type that represents the type of the type 
        ''' converter for this property. This type must derive from TypeConverter.</param> 
        Public Sub New(ByVal name As String, ByVal type As Type, ByVal category As String, ByVal description As String, ByVal defaultValue As Object, ByVal editor As String, _
        ByVal typeConverter As Type)
            Me.New(name, type.AssemblyQualifiedName, category, description, defaultValue, editor, _
            typeConverter.AssemblyQualifiedName)
        End Sub

        ''' <summary> 
        ''' Initializes a new instance of the PropertySpec class. 
        ''' </summary> 
        ''' <param name="name">The name of the property displayed in the property grid.</param> 
        ''' <param name="type">The fully qualified name of the type of the property.</param> 
        ''' <param name="category">The category under which the property is displayed in the 
        ''' property grid.</param> 
        ''' <param name="description">A string that is displayed in the help area of the 
        ''' property grid.</param> 
        ''' <param name="defaultValue">The default value of the property, or null if there is 
        ''' no default value.</param> 
        ''' <param name="editor">The Type that represents the type of the editor for this 
        ''' property. This type must derive from UITypeEditor.</param> 
        ''' <param name="typeConverter">The Type that represents the type of the type 
        ''' converter for this property. This type must derive from TypeConverter.</param> 
        Public Sub New(ByVal name As String, ByVal type As String, ByVal category As String, ByVal description As String, ByVal defaultValue As Object, ByVal editor As Type, _
        ByVal typeConverter As Type)
            Me.New(name, type, category, description, defaultValue, editor.AssemblyQualifiedName, _
            typeConverter.AssemblyQualifiedName)
        End Sub

        ''' <summary> 
        ''' Initializes a new instance of the PropertySpec class. 
        ''' </summary> 
        ''' <param name="name">The name of the property displayed in the property grid.</param> 
        ''' <param name="type">A Type that represents the type of the property.</param> 
        ''' <param name="category">The category under which the property is displayed in the 
        ''' property grid.</param> 
        ''' <param name="description">A string that is displayed in the help area of the 
        ''' property grid.</param> 
        ''' <param name="defaultValue">The default value of the property, or null if there is 
        ''' no default value.</param> 
        ''' <param name="editor">The Type that represents the type of the editor for this 
        ''' property. This type must derive from UITypeEditor.</param> 
        ''' <param name="typeConverter">The Type that represents the type of the type 
        ''' converter for this property. This type must derive from TypeConverter.</param> 
        Public Sub New(ByVal name As String, ByVal type As Type, ByVal category As String, ByVal description As String, ByVal defaultValue As Object, ByVal editor As Type, _
        ByVal typeConverter As Type)
            Me.New(name, type.AssemblyQualifiedName, category, description, defaultValue, editor.AssemblyQualifiedName, _
            typeConverter.AssemblyQualifiedName)
        End Sub

        ''' <summary> 
        ''' Gets or sets a collection of additional Attributes for this property. This can 
        ''' be used to specify attributes beyond those supported intrinsically by the 
        ''' PropertySpec class, such as ReadOnly and Browsable. 
        ''' </summary> 
        Public Property Attributes() As Attribute()
            Get
                Return m_attributes
            End Get
            Set(ByVal value As Attribute())
                m_attributes = value
            End Set
        End Property

        ''' <summary> 
        ''' Gets or sets the category name of this property. 
        ''' </summary> 
        Public Property Category() As String
            Get
                Return m_category
            End Get
            Set(ByVal value As String)
                m_category = value
            End Set
        End Property

        ''' <summary> 
        ''' Gets or sets the fully qualified name of the type converter 
        ''' type for this property. 
        ''' </summary> 
        Public Property ConverterTypeName() As String
            Get
                Return typeConverter
            End Get
            Set(ByVal value As String)
                typeConverter = value
            End Set
        End Property

        ''' <summary> 
        ''' Gets or sets the default value of this property. 
        ''' </summary> 
        Public Property DefaultValue() As Object
            Get
                Return m_defaultValue
            End Get
            Set(ByVal value As Object)
                m_defaultValue = value
            End Set
        End Property

        ''' <summary> 
        ''' Gets or sets the help text description of this property. 
        ''' </summary> 
        Public Property Description() As String
            Get
                Return m_description
            End Get
            Set(ByVal value As String)
                m_description = value
            End Set
        End Property

        ''' <summary> 
        ''' Gets or sets the fully qualified name of the editor type for 
        ''' this property. 
        ''' </summary> 
        Public Property EditorTypeName() As String
            Get
                Return editor
            End Get
            Set(ByVal value As String)
                editor = value
            End Set
        End Property

        ''' <summary> 
        ''' Gets or sets the name of this property. 
        ''' </summary> 
        Public Property Name() As String
            Get
                Return m_name
            End Get
            Set(ByVal value As String)
                m_name = value
            End Set
        End Property

        ''' <summary> 
        ''' Gets or sets the fully qualfied name of the type of this 
        ''' property. 
        ''' </summary> 
        Public Property TypeName() As String
            Get
                Return type
            End Get
            Set(ByVal value As String)
                type = value
            End Set
        End Property
    End Class

    ''' <summary> 
    ''' Provides data for the GetValue and SetValue events of the PropertyBag class. 
    ''' </summary> 
    Public Class PropertySpecEventArgs
        Inherits EventArgs
        Private m_property As PropertySpec
        Private val As Object

        ''' <summary> 
        ''' Initializes a new instance of the PropertySpecEventArgs class. 
        ''' </summary> 
        ''' <param name="property">The PropertySpec that represents the property whose 
        ''' value is being requested or set.</param> 
        ''' <param name="val">The current value of the property.</param> 
        Public Sub New(ByVal [property] As PropertySpec, ByVal val As Object)
            Me.m_property = [property]
            Me.val = val
        End Sub

        ''' <summary> 
        ''' Gets the PropertySpec that represents the property whose value is being 
        ''' requested or set. 
        ''' </summary> 
        Public ReadOnly Property [Property]() As PropertySpec
            Get
                Return m_property
            End Get
        End Property

        ''' <summary> 
        ''' Gets or sets the current value of the property. 
        ''' </summary> 
        Public Property Value() As Object
            Get
                Return val
            End Get
            Set(ByVal value As Object)
                val = value
            End Set
        End Property
    End Class

    ''' <summary> 
    ''' Represents the method that will handle the GetValue and SetValue events of the 
    ''' PropertyBag class. 
    ''' </summary> 
    Public Delegate Sub PropertySpecEventHandler(ByVal sender As Object, ByVal e As PropertySpecEventArgs)

    ''' <summary> 
    ''' Represents a collection of custom properties that can be selected into a 
    ''' PropertyGrid to provide functionality beyond that of the simple reflection 
    ''' normally used to query an object's properties. 
    ''' </summary> 
    Public Class PropertyBag
        Implements ICustomTypeDescriptor
#Region "PropertySpecCollection class definition"
        ''' <summary> 
        ''' Encapsulates a collection of PropertySpec objects. 
        ''' </summary> 
        <Serializable()> _
        Public Class PropertySpecCollection
            Implements IList
            Private innerArray As ArrayList

            ''' <summary> 
            ''' Initializes a new instance of the PropertySpecCollection class. 
            ''' </summary> 
            Public Sub New()
                innerArray = New ArrayList()
            End Sub

            ''' <summary> 
            ''' Gets the number of elements in the PropertySpecCollection. 
            ''' </summary> 
            ''' <value> 
            ''' The number of elements contained in the PropertySpecCollection. 
            ''' </value> 
            Public ReadOnly Property Count() As Integer Implements IList.Count
                Get
                    Return innerArray.Count
                End Get
            End Property

            ''' <summary> 
            ''' Gets a value indicating whether the PropertySpecCollection has a fixed size. 
            ''' </summary> 
            ''' <value> 
            ''' true if the PropertySpecCollection has a fixed size; otherwise, false. 
            ''' </value> 
            Public ReadOnly Property IsFixedSize() As Boolean Implements IList.IsFixedSize
                Get
                    Return False
                End Get
            End Property

            ''' <summary> 
            ''' Gets a value indicating whether the PropertySpecCollection is read-only. 
            ''' </summary> 
            Public ReadOnly Property IsReadOnly() As Boolean Implements IList.IsReadOnly
                Get
                    Return False
                End Get
            End Property

            ''' <summary> 
            ''' Gets a value indicating whether access to the collection is synchronized (thread-safe). 
            ''' </summary> 
            ''' <value> 
            ''' true if access to the PropertySpecCollection is synchronized (thread-safe); otherwise, false. 
            ''' </value> 
            Public ReadOnly Property IsSynchronized() As Boolean Implements IList.IsSynchronized
                Get
                    Return False
                End Get
            End Property

            ''' <summary> 
            ''' Gets an object that can be used to synchronize access to the collection. 
            ''' </summary> 
            ''' <value> 
            ''' An object that can be used to synchronize access to the collection. 
            ''' </value> 
            Private ReadOnly Property SyncRoot() As Object Implements ICollection.SyncRoot
                Get
                    Return Nothing
                End Get
            End Property

            ''' <summary> 
            ''' Gets or sets the element at the specified index. 
            ''' In C#, this property is the indexer for the PropertySpecCollection class. 
            ''' </summary> 
            ''' <param name="index">The zero-based index of the element to get or set.</param> 
            ''' <value> 
            ''' The element at the specified index. 
            ''' </value> 
            Default Public Property Item(ByVal index As Integer) As Object Implements IList.Item
                Get
                    Return DirectCast(innerArray(index), PropertySpec)
                End Get
                Set(ByVal value As Object)
                    innerArray(index) = value
                End Set
            End Property

            ''' <summary> 
            ''' Adds a PropertySpec to the end of the PropertySpecCollection. 
            ''' </summary> 
            ''' <param name="value">The PropertySpec to be added to the end of the PropertySpecCollection.</param> 
            ''' <returns>The PropertySpecCollection index at which the value has been added.</returns> 
            Public Function Add(ByVal value As PropertySpec) As Integer
                Dim index As Integer = innerArray.Add(value)

                Return index
            End Function

            ''' <summary> 
            ''' Adds the elements of an array of PropertySpec objects to the end of the PropertySpecCollection. 
            ''' </summary> 
            ''' <param name="array">The PropertySpec array whose elements should be added to the end of the 
            ''' PropertySpecCollection.</param> 
            Public Sub AddRange(ByVal array As PropertySpec())
                innerArray.AddRange(array)
            End Sub

            ''' <summary> 
            ''' Removes all elements from the PropertySpecCollection. 
            ''' </summary> 
            Public Sub Clear() Implements IList.Clear
                innerArray.Clear()
            End Sub

            ''' <summary> 
            ''' Determines whether a PropertySpec is in the PropertySpecCollection. 
            ''' </summary> 
            ''' <param name="item">The PropertySpec to locate in the PropertySpecCollection. The element to locate 
            ''' can be a null reference (Nothing in Visual Basic).</param> 
            ''' <returns>true if item is found in the PropertySpecCollection; otherwise, false.</returns> 
            Public Function Contains(ByVal item As PropertySpec) As Boolean
                Return innerArray.Contains(item)
            End Function

            ''' <summary> 
            ''' Determines whether a PropertySpec with the specified name is in the PropertySpecCollection. 
            ''' </summary> 
            ''' <param name="name">The name of the PropertySpec to locate in the PropertySpecCollection.</param> 
            ''' <returns>true if item is found in the PropertySpecCollection; otherwise, false.</returns> 
            Public Function Contains(ByVal name As String) As Boolean
                For Each spec As PropertySpec In innerArray
                    If spec.Name = name Then
                        Return True
                    End If
                Next

                Return False
            End Function

            ''' <summary> 
            ''' Copies the entire PropertySpecCollection to a compatible one-dimensional Array, starting at the 
            ''' beginning of the target array. 
            ''' </summary> 
            ''' <param name="array">The one-dimensional Array that is the destination of the elements copied 
            ''' from PropertySpecCollection. The Array must have zero-based indexing.</param> 
            Public Sub CopyTo(ByVal array As PropertySpec())
                innerArray.CopyTo(array)
            End Sub

            ''' <summary> 
            ''' Copies the PropertySpecCollection or a portion of it to a one-dimensional array. 
            ''' </summary> 
            ''' <param name="array">The one-dimensional Array that is the destination of the elements copied 
            ''' from the collection.</param> 
            ''' <param name="index">The zero-based index in array at which copying begins.</param> 
            Public Sub CopyTo(ByVal array As PropertySpec(), ByVal index As Integer)
                innerArray.CopyTo(array, index)
            End Sub

            ''' <summary> 
            ''' Returns an enumerator that can iterate through the PropertySpecCollection. 
            ''' </summary> 
            ''' <returns>An IEnumerator for the entire PropertySpecCollection.</returns> 
            Public Function GetEnumerator() As IEnumerator Implements IList.GetEnumerator
                Return innerArray.GetEnumerator()
            End Function

            ''' <summary> 
            ''' Searches for the specified PropertySpec and returns the zero-based index of the first 
            ''' occurrence within the entire PropertySpecCollection. 
            ''' </summary> 
            ''' <param name="value">The PropertySpec to locate in the PropertySpecCollection.</param> 
            ''' <returns>The zero-based index of the first occurrence of value within the entire PropertySpecCollection, 
            ''' if found; otherwise, -1.</returns> 
            Public Function IndexOf(ByVal value As PropertySpec) As Integer
                Return innerArray.IndexOf(value)
            End Function

            ''' <summary> 
            ''' Searches for the PropertySpec with the specified name and returns the zero-based index of 
            ''' the first occurrence within the entire PropertySpecCollection. 
            ''' </summary> 
            ''' <param name="name">The name of the PropertySpec to locate in the PropertySpecCollection.</param> 
            ''' <returns>The zero-based index of the first occurrence of value within the entire PropertySpecCollection, 
            ''' if found; otherwise, -1.</returns> 
            Public Function IndexOf(ByVal name As String) As Integer
                Dim i As Integer = 0

                For Each spec As PropertySpec In innerArray
                    If spec.Name = name Then
                        Return i
                    End If

                    i += 1
                Next

                Return -1
            End Function

            ''' <summary> 
            ''' Inserts a PropertySpec object into the PropertySpecCollection at the specified index. 
            ''' </summary> 
            ''' <param name="index">The zero-based index at which value should be inserted.</param> 
            ''' <param name="value">The PropertySpec to insert.</param> 
            Public Sub Insert(ByVal index As Integer, ByVal value As PropertySpec)
                innerArray.Insert(index, value)
            End Sub

            ''' <summary> 
            ''' Removes the first occurrence of a specific object from the PropertySpecCollection. 
            ''' </summary> 
            ''' <param name="obj">The PropertySpec to remove from the PropertySpecCollection.</param> 
            Public Sub Remove(ByVal obj As PropertySpec)
                innerArray.Remove(obj)
            End Sub

            ''' <summary> 
            ''' Removes the property with the specified name from the PropertySpecCollection. 
            ''' </summary> 
            ''' <param name="name">The name of the PropertySpec to remove from the PropertySpecCollection.</param> 
            Public Sub Remove(ByVal name As String)
                Dim index As Integer = IndexOf(name)
                RemoveAt(index)
            End Sub

            ''' <summary> 
            ''' Removes the object at the specified index of the PropertySpecCollection. 
            ''' </summary> 
            ''' <param name="index">The zero-based index of the element to remove.</param> 
            Public Sub RemoveAt(ByVal index As Integer) Implements IList.RemoveAt
                innerArray.RemoveAt(index)
            End Sub

            ''' <summary> 
            ''' Copies the elements of the PropertySpecCollection to a new PropertySpec array. 
            ''' </summary> 
            ''' <returns>A PropertySpec array containing copies of the elements of the PropertySpecCollection.</returns> 
            Public Function ToArray() As PropertySpec()
                Return DirectCast(innerArray.ToArray(GetType(PropertySpec)), PropertySpec())
            End Function

#Region "Explicit interface implementations for ICollection and IList"
            ''' <summary> 
            ''' This member supports the .NET Framework infrastructure and is not intended to be used directly from your code. 
            ''' </summary> 
            Private Sub CopyTo(ByVal array As Array, ByVal index As Integer) Implements ICollection.CopyTo
                CopyTo(DirectCast(array, PropertySpec()), index)
            End Sub

            ''' <summary> 
            ''' This member supports the .NET Framework infrastructure and is not intended to be used directly from your code. 
            ''' </summary> 
            Private Function Add(ByVal value As Object) As Integer Implements IList.Add
                Return Add(DirectCast(value, PropertySpec))
            End Function

            ''' <summary> 
            ''' This member supports the .NET Framework infrastructure and is not intended to be used directly from your code. 
            ''' </summary> 
            Private Function Contains(ByVal obj As Object) As Boolean Implements IList.Contains
                Return Contains(DirectCast(obj, PropertySpec))
            End Function

            '''' <summary> 
            '''' This member supports the .NET Framework infrastructure and is not intended to be used directly from your code. 
            '''' </summary> 
            'Default Property Item(ByVal index As Integer) As Object Implements IList.this
            '    Get
            '        Return DirectCast(Me, PropertySpecCollection)(index)
            '    End Get
            '    Set(ByVal value As Object)
            '        DirectCast(Me, PropertySpecCollection)(index) = DirectCast(value, PropertySpec)
            '    End Set
            'End Property

            ''' <summary> 
            ''' This member supports the .NET Framework infrastructure and is not intended to be used directly from your code. 
            ''' </summary> 
            Private Function IndexOf(ByVal obj As Object) As Integer Implements IList.IndexOf
                Return IndexOf(DirectCast(obj, PropertySpec))
            End Function

            ''' <summary> 
            ''' This member supports the .NET Framework infrastructure and is not intended to be used directly from your code. 
            ''' </summary> 
            Private Sub Insert(ByVal index As Integer, ByVal value As Object) Implements IList.Insert
                Insert(index, DirectCast(value, PropertySpec))
            End Sub

            ''' <summary> 
            ''' This member supports the .NET Framework infrastructure and is not intended to be used directly from your code. 
            ''' </summary> 
            Private Sub Remove(ByVal value As Object) Implements IList.Remove
                Remove(DirectCast(value, PropertySpec))
            End Sub
#End Region
        End Class
#End Region
#Region "PropertySpecDescriptor class definition"
        Private Class PropertySpecDescriptor
            Inherits PropertyDescriptor
            Private bag As PropertyBag
            Private item As PropertySpec

            Public Sub New(ByVal item As PropertySpec, ByVal bag As PropertyBag, ByVal name As String, ByVal attrs As Attribute())
                MyBase.New(name, attrs)
                Me.bag = bag
                Me.item = item
            End Sub

            Public Overloads Overrides ReadOnly Property ComponentType() As Type
                Get
                    Return item.[GetType]()
                End Get
            End Property

            Public Overloads Overrides ReadOnly Property IsReadOnly() As Boolean
                Get
                    Return (Attributes.Matches(ReadOnlyAttribute.Yes))
                End Get
            End Property

            Public Overloads Overrides ReadOnly Property PropertyType() As Type
                Get
                    Return Type.[GetType](item.TypeName)
                End Get
            End Property

            Public Overloads Overrides Function CanResetValue(ByVal component As Object) As Boolean
                If item.DefaultValue Is Nothing Then
                    Return False
                Else
                    Return Not Me.GetValue(component).Equals(item.DefaultValue)
                End If
            End Function

            Public Overloads Overrides Function GetValue(ByVal component As Object) As Object
                ' Have the property bag raise an event to get the current value 
                ' of the property. 

                Dim e As New PropertySpecEventArgs(item, Nothing)
                bag.OnGetValue(e)
                Return e.Value
            End Function

            Public Overloads Overrides Sub ResetValue(ByVal component As Object)
                SetValue(component, item.DefaultValue)
            End Sub

            Public Overloads Overrides Sub SetValue(ByVal component As Object, ByVal value As Object)
                ' Have the property bag raise an event to set the current value 
                ' of the property. 

                Dim e As New PropertySpecEventArgs(item, value)
                bag.OnSetValue(e)
            End Sub

            Public Overloads Overrides Function ShouldSerializeValue(ByVal component As Object) As Boolean
                Dim val As Object = Me.GetValue(component)

                If item.DefaultValue Is Nothing AndAlso val Is Nothing Then
                    Return False
                Else
                    Return Not val.Equals(item.DefaultValue)
                End If
            End Function
        End Class
#End Region

        Private m_defaultProperty As String
        Private m_properties As PropertySpecCollection

        ''' <summary> 
        ''' Initializes a new instance of the PropertyBag class. 
        ''' </summary> 
        Public Sub New()
            m_defaultProperty = Nothing
            m_properties = New PropertySpecCollection()
        End Sub

        ''' <summary> 
        ''' Gets or sets the name of the default property in the collection. 
        ''' </summary> 
        Public Property DefaultProperty() As String
            Get
                Return m_defaultProperty
            End Get
            Set(ByVal value As String)
                m_defaultProperty = value
            End Set
        End Property

        ''' <summary> 
        ''' Gets the collection of properties contained within this PropertyBag. 
        ''' </summary> 
        Public ReadOnly Property Properties() As PropertySpecCollection
            Get
                Return m_properties
            End Get
        End Property

        ''' <summary> 
        ''' Occurs when a PropertyGrid requests the value of a property. 
        ''' </summary> 
        Public Event GetValue As PropertySpecEventHandler

        ''' <summary> 
        ''' Occurs when the user changes the value of a property in a PropertyGrid. 
        ''' </summary> 
        Public Event SetValue As PropertySpecEventHandler

        ''' <summary> 
        ''' Raises the GetValue event. 
        ''' </summary> 
        ''' <param name="e">A PropertySpecEventArgs that contains the event data.</param> 
        Protected Overridable Sub OnGetValue(ByVal e As PropertySpecEventArgs)
            RaiseEvent GetValue(Me, e)
        End Sub

        ''' <summary> 
        ''' Raises the SetValue event. 
        ''' </summary> 
        ''' <param name="e">A PropertySpecEventArgs that contains the event data.</param> 
        Protected Overridable Sub OnSetValue(ByVal e As PropertySpecEventArgs)
            RaiseEvent SetValue(Me, e)
        End Sub

#Region "ICustomTypeDescriptor explicit interface definitions"
        ' Most of the functions required by the ICustomTypeDescriptor are 
        ' merely pssed on to the default TypeDescriptor for this type, 
        ' which will do something appropriate. The exceptions are noted 
        ' below. 
        Private Function GetAttributes() As AttributeCollection Implements ICustomTypeDescriptor.GetAttributes
            Return TypeDescriptor.GetAttributes(Me, True)
        End Function

        Private Function GetClassName() As String Implements ICustomTypeDescriptor.GetClassName
            Return TypeDescriptor.GetClassName(Me, True)
        End Function

        Private Function GetComponentName() As String Implements ICustomTypeDescriptor.GetComponentName
            Return TypeDescriptor.GetComponentName(Me, True)
        End Function

        Private Function GetConverter() As TypeConverter Implements ICustomTypeDescriptor.GetConverter
            Return TypeDescriptor.GetConverter(Me, True)
        End Function

        Private Function GetDefaultEvent() As EventDescriptor Implements ICustomTypeDescriptor.GetDefaultEvent
            Return TypeDescriptor.GetDefaultEvent(Me, True)
        End Function

        Private Function GetDefaultProperty() As PropertyDescriptor Implements ICustomTypeDescriptor.GetDefaultProperty
            ' This function searches the property list for the property 
            ' with the same name as the DefaultProperty specified, and 
            ' returns a property descriptor for it. If no property is 
            ' found that matches DefaultProperty, a null reference is 
            ' returned instead. 

            Dim propertySpec As PropertySpec = Nothing
            If m_defaultProperty IsNot Nothing Then
                Dim index As Integer = m_properties.IndexOf(m_defaultProperty)
                propertySpec = m_properties(index)
            End If

            If propertySpec IsNot Nothing Then
                Return New PropertySpecDescriptor(propertySpec, Me, propertySpec.Name, Nothing)
            Else
                Return Nothing
            End If
        End Function

        Private Function GetEditor(ByVal editorBaseType As Type) As Object Implements ICustomTypeDescriptor.GetEditor
            Return TypeDescriptor.GetEditor(Me, editorBaseType, True)
        End Function

        Private Function GetEvents() As EventDescriptorCollection Implements ICustomTypeDescriptor.GetEvents
            Return TypeDescriptor.GetEvents(Me, True)
        End Function

        Private Function GetEvents(ByVal attributes As Attribute()) As EventDescriptorCollection Implements ICustomTypeDescriptor.GetEvents
            Return TypeDescriptor.GetEvents(Me, attributes, True)
        End Function

        Private Function GetProperties() As PropertyDescriptorCollection Implements ICustomTypeDescriptor.GetProperties
            Return DirectCast(Me, ICustomTypeDescriptor).GetProperties(New Attribute(-1) {})
        End Function

        Private Function GetProperties(ByVal attributes As Attribute()) As PropertyDescriptorCollection Implements ICustomTypeDescriptor.GetProperties
            ' Rather than passing this function on to the default TypeDescriptor, 
            ' which would return the actual properties of PropertyBag, I construct 
            ' a list here that contains property descriptors for the elements of the 
            ' Properties list in the bag. 

            Dim props As New ArrayList()

            For Each [property] As PropertySpec In m_properties
                Dim attrs As New ArrayList()

                ' If a category, description, editor, or type converter are specified 
                ' in the PropertySpec, create attributes to define that relationship. 
                If [property].Category IsNot Nothing Then
                    attrs.Add(New CategoryAttribute([property].Category))
                End If

                If [property].Description IsNot Nothing Then
                    attrs.Add(New DescriptionAttribute([property].Description))
                End If

                If [property].EditorTypeName IsNot Nothing Then
                    attrs.Add(New EditorAttribute([property].EditorTypeName, GetType(UITypeEditor)))
                End If

                If [property].ConverterTypeName IsNot Nothing Then
                    attrs.Add(New TypeConverterAttribute([property].ConverterTypeName))
                End If

                ' Additionally, append the custom attributes associated with the 
                ' PropertySpec, if any. 
                If [property].Attributes IsNot Nothing Then
                    attrs.AddRange([property].Attributes)
                End If

                Dim attrArray As Attribute() = DirectCast(attrs.ToArray(GetType(Attribute)), Attribute())

                ' Create a new property descriptor for the property item, and add 
                ' it to the list. 
                Dim pd As New PropertySpecDescriptor([property], Me, [property].Name, attrArray)
                props.Add(pd)
            Next

            ' Convert the list of PropertyDescriptors to a collection that the 
            ' ICustomTypeDescriptor can use, and return it. 
            Dim propArray As PropertyDescriptor() = DirectCast(props.ToArray(GetType(PropertyDescriptor)), PropertyDescriptor())
            Return New PropertyDescriptorCollection(propArray)
        End Function

        Private Function GetPropertyOwner(ByVal pd As PropertyDescriptor) As Object Implements ICustomTypeDescriptor.GetPropertyOwner
            Return Me
        End Function
#End Region
    End Class

    ''' <summary> 
    ''' An extension of PropertyBag that manages a table of property values, in 
    ''' addition to firing events when property values are requested or set. 
    ''' </summary> 
    Public Class PropertyTable
        Inherits PropertyBag
        Private propValues As Hashtable

        ''' <summary> 
        ''' Initializes a new instance of the PropertyTable class. 
        ''' </summary> 
        Public Sub New()
            propValues = New Hashtable()
        End Sub

        ''' <summary> 
        ''' Gets or sets the value of the property with the specified name. 
        ''' <p>In C#, this property is the indexer of the PropertyTable class.</p> 
        ''' </summary> 
        Default Public Property Item(ByVal key As String) As Object
            Get
                Return propValues(key)
            End Get
            Set(ByVal value As Object)
                propValues(key) = value
            End Set
        End Property

        ''' <summary> 
        ''' This member overrides PropertyBag.OnGetValue. 
        ''' </summary> 
        Protected Overloads Overrides Sub OnGetValue(ByVal e As PropertySpecEventArgs)
            e.Value = propValues(e.[Property].Name)
            MyBase.OnGetValue(e)
        End Sub

        ''' <summary> 
        ''' This member overrides PropertyBag.OnSetValue. 
        ''' </summary> 
        Protected Overloads Overrides Sub OnSetValue(ByVal e As PropertySpecEventArgs)
            propValues(e.[Property].Name) = e.Value
            MyBase.OnSetValue(e)
        End Sub
    End Class
End Namespace