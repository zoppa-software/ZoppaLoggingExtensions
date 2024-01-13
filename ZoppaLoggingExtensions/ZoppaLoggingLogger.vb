Option Strict On
Option Explicit On

Imports System.Threading
Imports Microsoft.Extensions.Logging

''' <summary>ログ出力クラスです。</summary>
Public NotInheritable Class ZoppaLoggingLogger
    Implements ILogger

    ' ミューティックス
    Private ReadOnly _mutex As Mutex

    ' 内部ロガー
    Private _logger As LocalLogger

    ''' <summary>カテゴリ名です。</summary>
    Public ReadOnly CategoryName As String

    ''' <summary>ログ設定です。</summary>
    Public ReadOnly Config As Func(Of ZoppaLoggingConfiguration)

    ''' <summary>スコープリストです。</summary>
    Public ReadOnly Scopes As New List(Of ScopeBase)

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="categoryName">カテゴリ名。</param>
    ''' <param name="config">ログ設定。</param>
    Public Sub New(categoryName As String, config As Func(Of ZoppaLoggingConfiguration))
        Me.CategoryName = categoryName
        Me.Config = config

        Me._mutex = New Mutex(False, $"Global\zoppalog_{categoryName}")
    End Sub

    ''' <summary>スコープを開始します。</summary>
    ''' <typeparam name="TState">ステート型。</typeparam>
    ''' <param name="state">ステート。</param>
    ''' <returns>スコープ。</returns>
    Public Function BeginScope(Of TState)(state As TState) As IDisposable Implements ILogger.BeginScope
        Return New Scope(Of TState)(Me.Scopes, state)
    End Function

    ''' <summary>指定されたログレベルが有効かどうかを返します。</summary>
    ''' <param name="logLevel">ログレベル。</param>
    ''' <returns>有効ならば真。</returns>
    Public Function IsEnabled(logLevel As LogLevel) As Boolean Implements ILogger.IsEnabled
        Return (logLevel >= Me.Config().MinimumLogLevel)
    End Function

    ''' <summary>ログ出力します。</summary>
    ''' <typeparam name="TState">ステート型。</typeparam>
    ''' <param name="logLevel">ログレベル。</param>
    ''' <param name="eventId">イベントID。</param>
    ''' <param name="state">ステート。</param>
    ''' <param name="exception">例外。</param>
    ''' <param name="formatter">フォーマッター。</param>
    Public Sub Log(Of TState)(logLevel As LogLevel,
                              eventId As EventId,
                              state As TState,
                              exception As Exception,
                              formatter As Func(Of TState, Exception, String)) Implements ILogger.Log
        If Me._logger Is Nothing Then
            Me._mutex.WaitOne()
            Try
                Me._logger = LocalLogger.Create(Me.CategoryName, Me.Config)
            Finally
                Me._mutex.ReleaseMutex()
            End Try
        End If

        If Me.IsEnabled(logLevel) Then
            Me._logger.Stock(New LocalLogger.LogData(Me.CategoryName, logLevel, eventId, formatter(state, exception), Me.Scopes))
            Me._logger.Write(Me._mutex)
        End If
    End Sub

    ''' <summary>ログ出力終了を待機します。</summary>
    Public Sub WaitFinish()
        Me._logger?.WaitFinish()
    End Sub

    ''' <summary>スコープ基底クラスです。</summary>
    Public MustInherit Class ScopeBase
        Implements IDisposable

        ' スコープリスト
        Private ReadOnly _scopes As List(Of ScopeBase)

        ''' <summary>スコープ名を取得します。</summary>
        Public MustOverride ReadOnly Property ScopeName As String

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="scopes">スコープリスト。</param>
        Protected Sub New(scopes As List(Of ScopeBase))
            Me._scopes = scopes
            Me._scopes.Add(Me)
        End Sub

        ''' <summary>スコープを終了します。</summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            Me._scopes.Remove(Me)
        End Sub

    End Class

    ''' <summary>スコープクラスです。</summary>
    ''' <typeparam name="TState">ステート型。</typeparam>
    Private NotInheritable Class Scope(Of TState)
        Inherits ScopeBase

        ' ステート
        Private ReadOnly _state As TState

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="scopes">スコープリスト。</param>
        ''' <param name="state">ステート。</param>
        Public Sub New(scopes As List(Of ScopeBase), state As TState)
            MyBase.New(scopes)
            Me._state = state
        End Sub

        ''' <summary>スコープ名を取得します。</summary>
        Public Overrides ReadOnly Property ScopeName As String
            Get
                Return Me._state.ToString()
            End Get
        End Property

    End Class

End Class
