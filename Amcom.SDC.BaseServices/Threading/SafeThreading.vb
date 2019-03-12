Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Threading
Imports System.Reflection

Namespace SDCThreading
    ''' <summary> 
    ''' simple no-arg delegate type; can use this for anonymous methods, e.g. 
    ''' <code> 
    ''' SafeThread safeThrd = new SafeThread((SimpleDelegate) delegate { dosomething(); }); 
    ''' </code> 
    ''' </summary> 
    Public Delegate Sub SimpleDelegate()

    ''' <summary> 
    ''' delegate for thread-threw-exception event 
    ''' </summary> 
    ''' <param name="thrd">the SafeThread that threw the exception</param> 
    ''' <param name="ex">the exception throws</param> 
    Public Delegate Sub ThreadThrewExceptionHandler(ByVal thrd As SafeThread, ByVal ex As Exception)

    ''' <summary> 
    ''' delegate for thread-completed event 
    ''' </summary> 
    ''' <param name="thrd">the SafeThread that completed processing</param> 
    ''' <param name="hadException">true if the thread terminated due to an exception</param> 
    ''' <param name="ex">the exception that terminated the thread, or null if completed successfully</param> 
    Public Delegate Sub ThreadCompletedHandler(ByVal thrd As SafeThread, ByVal hadException As Boolean, ByVal ex As Exception)

    ''' <summary> 
    ''' This class implements a Thread wrapper to trap unhandled exceptions 
    ''' thrown by the thread-start delegate. Add ThreadException event 
    ''' handlers to be notified of such exceptions and take custom actions 
    ''' (such as restart, clean-up, et al, depending on what the SafeThread was 
    ''' doing in your application). Add ThreadCompleted event handlers to be 
    ''' notified when the thread has completed processing. 
    ''' </summary> 
    Partial Public Class SafeThread
        Inherits MarshalByRefObject
        ''' <summary> 
        ''' internal thread 
        ''' </summary> 
        Protected _thread As Thread
        ''' <summary> 
        ''' gets the internal thread being used 
        ''' </summary> 
        Public ReadOnly Property ThreadObject() As Thread
            Get
                Return _thread
            End Get
        End Property

        ''' <summary> 
        ''' the thread-start object, if any 
        ''' </summary> 
        Protected _Ts As ThreadStart
        ''' <summary> 
        ''' the parameterized thread-start object, if any 
        ''' </summary> 
        Protected _Pts As ParameterizedThreadStart
        ''' <summary> 
        ''' the SimpleDelegate target, if any 
        ''' </summary> 
        Protected dlg As SimpleDelegate

        ''' <summary> 
        ''' the thread-start argument object, if any 
        ''' </summary> 
        Protected _arg As Object
        ''' <summary> 
        ''' gets the thread-start argument, if any 
        ''' </summary> 
        Public ReadOnly Property ThreadStartArg() As Object
            Get
                Return _arg
            End Get
        End Property
        ''' <summary> 
        ''' the last exception thrown 
        ''' </summary> 
        Protected _lastException As Exception
        ''' <summary> 
        ''' gets the last exception thrown 
        ''' </summary> 
        Public ReadOnly Property LastException() As Exception
            Get
                Return _lastException
            End Get
        End Property
        ''' <summary> 
        ''' the name of the internal thread 
        ''' </summary> 
        Private _name As String
        ''' <summary> 
        ''' gets/sets the name of the internal thread 
        ''' </summary> 
        Public Property Name() As String
            Get
                If _name Is Nothing Then
                    Return "SafeThread#" + Me.GetHashCode().ToString()
                End If
                Return _name
            End Get
            Set(ByVal value As String)
                _name = value
            End Set
        End Property

        Private _tag As Object
        ''' <summary> 
        ''' object tag - use to hold extra info about the SafeThread 
        ''' </summary> 
        Public Property Tag() As Object
            Get
                Return _tag
            End Get
            Set(ByVal value As Object)
                _tag = value
            End Set
        End Property

        ''' <summary> 
        ''' default constructor for SafeThread 
        ''' </summary> 
        Public Sub New()
            MyBase.New()
        End Sub

        ''' <summary> 
        ''' SafeThread constructor using ThreadStart object 
        ''' </summary> 
        ''' <param name="ts">ThreadStart object to use</param> 
        Public Sub New(ByVal ts As ThreadStart)
            Me.New()
            _Ts = ts
            _thread = New Thread(ts)
        End Sub

        ''' <summary> 
        ''' SafeThread constructor using ParameterizedThreadStart object 
        ''' </summary> 
        ''' <param name="pts">ParameterizedThreadStart to use</param> 
        Public Sub New(ByVal pts As ParameterizedThreadStart)
            Me.New()
            _Pts = pts
            _thread = New Thread(pts)
        End Sub

        ''' <summary> 
        ''' SafeThread constructor using SimpleDelegate object for anonymous methods, e.g. 
        ''' <code> 
        ''' SafeThread safeThrd = new SafeThread((SimpleDelegate) delegate { dosomething(); }); 
        ''' </code> 
        ''' </summary> 
        ''' <param name="sd"></param> 
        Public Sub New(ByVal sd As SimpleDelegate)
            Me.New()
            dlg = sd
            _Pts = New ParameterizedThreadStart(AddressOf Me.CallDelegate)
            _thread = New Thread(_Pts)
        End Sub

        ''' <summary> 
        ''' thread-threw-exception event 
        ''' </summary> 
        Public Event ThreadException As ThreadThrewExceptionHandler

        ''' <summary> 
        ''' called when a thread throws an exception 
        ''' </summary> 
        ''' <param name="ex">Exception thrown</param> 
        Protected Sub OnThreadException(ByVal ex As Exception)
            Try
                If TypeOf ex Is ThreadAbortException AndAlso Not ShouldReportThreadAbort Then
                    Return
                End If

                RaiseEvent ThreadException(Me, ex)
            Catch
            End Try
        End Sub

        ''' <summary> 
        ''' thread-completed event 
        ''' </summary> 
        Public Event ThreadCompleted As ThreadCompletedHandler

        ''' <summary> 
        ''' called when a thread completes processing 
        ''' </summary> 
        Protected Sub OnThreadCompleted(ByVal bHadException As Boolean, ByVal ex As Exception)
            Try
                RaiseEvent ThreadCompleted(Me, bHadException, ex)
            Catch
            End Try
        End Sub

        ''' <summary> 
        ''' starts thread with target if any 
        ''' </summary> 
        Protected Sub startTarget()
            Dim exceptn As Exception = Nothing
            Dim bHadException As Boolean = False
            Try
                bThreadIsAborting = False
                If _Ts IsNot Nothing Then
                    _Ts.Invoke()
                ElseIf _Pts IsNot Nothing Then
                    _Pts.Invoke(_arg)
                End If
            Catch ex As Exception
                Dim st As New StackTrace(True)
                Dim sf As StackFrame = st.GetFrame(1)
                bHadException = True
                exceptn = ex
                Me._lastException = ex

                OnThreadException(ex)
            Finally
                OnThreadCompleted(bHadException, exceptn)
            End Try
        End Sub

        ''' <summary> 
        ''' thread-start internal method for SimpleDelegate target 
        ''' </summary> 
        ''' <param name="arg">unused</param> 
        Protected Sub CallDelegate(ByVal arg As Object)
            Me.dlg.Invoke()
        End Sub

        ''' <summary> 
        ''' starts thread execution 
        ''' </summary> 
        Public Sub Start()
            _thread = New Thread(New ThreadStart(AddressOf startTarget))
            _thread.Name = Me.Name
            If _aptState IsNot Nothing Then
                _thread.TrySetApartmentState(DirectCast(_aptState, ApartmentState))
            End If
            _thread.Start()
        End Sub

        ''' <summary> 
        ''' starts thread execution with parameter 
        ''' </summary> 
        ''' <param name="val">parameter object</param> 
        Public Sub Start(ByVal val As Object)
            _arg = val
            Start()
        End Sub

        ''' <summary> 
        ''' flag to control whether thread-abort exception is reported or not 
        ''' </summary> 
        Protected bReportThreadAbort As Boolean = False
        ''' <summary> 
        ''' gets/sets a flag to control whether thread-abort exception is reported or not 
        ''' </summary> 
        Public Property ShouldReportThreadAbort() As Boolean
            Get
                Return bReportThreadAbort
            End Get
            Set(ByVal value As Boolean)
                bReportThreadAbort = value
            End Set
        End Property

        ''' <summary> 
        ''' flag for when thread is aborting 
        ''' </summary> 
        Protected bThreadIsAborting As Boolean = False

        ''' <summary> 
        ''' abort the thread execution 
        ''' </summary> 
        Public Sub Abort()
            bThreadIsAborting = True
            _thread.Abort()
            If bReportThreadAbort Then
                RaiseEvent ThreadException(Me, New Exception("Thread Aborted"))
            End If
        End Sub

        ''' <summary> 
        ''' gets or sets the Culture for the current thread. 
        ''' </summary> 
        Public ReadOnly Property CurrentCulture() As System.Globalization.CultureInfo
            Get
                If _thread IsNot Nothing Then
                    Return _thread.CurrentCulture
                End If
                Return Nothing
            End Get
        End Property

        ''' <summary> 
        ''' gets or sets the current culture used by the Resource Manager 
        ''' to look up culture-specific resources at run time. 
        ''' </summary> 
        Public ReadOnly Property CurrentUICulture() As System.Globalization.CultureInfo
            Get
                If _thread IsNot Nothing Then
                    Return _thread.CurrentUICulture
                End If
                Return Nothing
            End Get
        End Property

        ''' <summary> 
        ''' gets an System.Threading.ExecutionContext object that contains information 
        ''' about the various contexts of the current thread. 
        ''' </summary> 
        Public ReadOnly Property ExecutionContext() As ExecutionContext
            Get
                If _thread IsNot Nothing Then
                    Return _thread.ExecutionContext
                End If
                Return Nothing
            End Get
        End Property
        ''' <summary> 
        ''' Returns an System.Threading.ApartmentState value indicating the apparent state. 
        ''' </summary> 
        ''' <returns></returns> 
        Public Function GetApartmentState() As ApartmentState
            If _thread IsNot Nothing Then
                Return _thread.GetApartmentState()
            End If
            Return ApartmentState.Unknown
        End Function

        ''' <summary> 
        ''' Interrupts a thread that is in the WaitSleepJoin thread state. 
        ''' </summary> 
        Public Sub Interrupt()
            If _thread IsNot Nothing Then
                _thread.Interrupt()
            End If
        End Sub

        ''' <summary> 
        ''' gets a value indicating the execution status of the thread 
        ''' </summary> 
        Public ReadOnly Property IsAlive() As Boolean
            Get
                If _thread IsNot Nothing Then
                    Return _thread.IsAlive
                End If
                Return False
            End Get
        End Property

        ''' <summary> 
        ''' Gets or sets a value indicating whether or not a thread is a background thread 
        ''' </summary> 
        Public Property IsBackground() As Boolean
            Get
                If _thread IsNot Nothing Then
                    Return _thread.IsBackground
                End If
                Return False
            End Get
            Set(ByVal value As Boolean)
                If _thread IsNot Nothing Then
                    _thread.IsBackground = value
                End If
            End Set
        End Property

        ''' <summary> 
        ''' gets a value indicating whether or not a thread belongs to the managed thread pool 
        ''' </summary> 
        Public ReadOnly Property IsThreadPoolThread() As Boolean
            Get
                If _thread IsNot Nothing Then
                    Return _thread.IsThreadPoolThread
                End If
                Return False
            End Get
        End Property

        ''' <summary> 
        ''' Blocks the calling thread until a thread terminates, 
        ''' while continuing to perform standard COM and SendMessage pumping. 
        ''' </summary> 
        Public Sub Join()
            If _thread IsNot Nothing Then
                _thread.Join()
            End If
        End Sub

        ''' <summary> 
        ''' Blocks the calling thread until a thread terminates or the specified time elapses, 
        ''' while continuing to perform standard COM and SendMessage pumping. 
        ''' </summary> 
        ''' <param name="millisecondsTimeout">the number of milliseconds to wait for the 
        ''' thread to terminate</param> 
        Public Function Join(ByVal millisecondsTimeout As Integer) As Boolean
            If _thread IsNot Nothing Then
                Return _thread.Join(millisecondsTimeout)
            End If
            Return False
        End Function

        ''' <summary> 
        ''' Blocks the calling thread until a thread terminates or the specified time elapses, 
        ''' while continuing to perform standard COM and SendMessage pumping. 
        ''' </summary> 
        ''' <param name="timeout">a System.TimeSpan set to the amount of time to wait 
        ''' for the thread to terminate </param> 
        Public Function Join(ByVal timeout As TimeSpan) As Boolean
            If _thread IsNot Nothing Then
                Return _thread.Join(timeout)
            End If
            Return False
        End Function

        ''' <summary> 
        ''' Gets a unique identifier for the current managed thread 
        ''' </summary> 
        Public ReadOnly Property ManagedThreadId() As Integer
            Get
                If _thread IsNot Nothing Then
                    Return _thread.ManagedThreadId
                End If
                Return 0
            End Get
        End Property

        ''' <summary> 
        ''' gets or sets a value indicating the scheduling priority of a thread 
        ''' </summary> 
        Public Property Priority() As ThreadPriority
            Get
                If _thread IsNot Nothing Then
                    Return _thread.Priority
                End If
                Return ThreadPriority.Lowest
            End Get
            Set(ByVal value As ThreadPriority)
                If _thread IsNot Nothing Then
                    _thread.Priority = value
                End If
            End Set
        End Property

        Private _aptState As Object
        ''' <summary> 
        ''' sets the ApartmentState of a thread before it is started 
        ''' </summary> 
        ''' <param name="state">ApartmentState</param> 
        Public Sub SetApartmentState(ByVal state As ApartmentState)
            If _thread IsNot Nothing Then
                _thread.SetApartmentState(state)
            Else
                _aptState = state
            End If
        End Sub

        ''' <summary> 
        ''' gets a value containing the states of the current thread 
        ''' </summary> 
        Public ReadOnly Property ThreadState() As ThreadState
            Get
                If _thread IsNot Nothing Then
                    Return _thread.ThreadState
                End If
                Return ThreadState.Unstarted
            End Get
        End Property

        ''' <summary> 
        ''' returns a System.String that represents the current System.Object 
        ''' </summary> 
        ''' <returns></returns> 
        Public Overloads Overrides Function ToString() As String
            If _thread IsNot Nothing Then
                Return _thread.ToString()
            End If
            Return MyBase.ToString()
        End Function

        ''' <summary> 
        ''' sets the ApartmentState of a thread before it is started 
        ''' </summary> 
        ''' <param name="state">ApartmentState</param> 
        Public Function TrySetApartmentState(ByVal state As ApartmentState) As Boolean
            If _thread IsNot Nothing Then
                Return _thread.TrySetApartmentState(state)
            End If
            _aptState = state
            Return False
        End Function
    End Class
End Namespace