Option Strict On
Option Explicit On

Imports System.Runtime.CompilerServices
Imports Microsoft.Extensions.DependencyInjection
Imports Microsoft.Extensions.DependencyInjection.Extensions
Imports Microsoft.Extensions.Logging
Imports Microsoft.Extensions.Logging.Configuration

''' <summary>ログ出力拡張クラスです。</summary>
Public Module ZoppaLogProcessExtensions

    ''' <summary>サービスの登録を行います。</summary>
    ''' <param name="builder">ログビルダー。</param>
    ''' <returns>ログビルダー。</returns>
    <Extension()>
    Public Function AddZoppaLogging(builder As ILoggingBuilder) As ILoggingBuilder
        builder.AddConfiguration()

        builder.Services.TryAddEnumerable(
            ServiceDescriptor.Singleton(Of ILoggerProvider, ZoppaLoggingProvider)()
        )

        LoggerProviderOptions.RegisterProviderOptions(Of ZoppaLoggingConfiguration, ZoppaLoggingProvider)(builder.Services)

        Return builder
    End Function

    ''' <summary>サービスの登録を行います。</summary>
    ''' <param name="builder">ログビルダー。</param>
    ''' <param name="configure">ログ設定。</param>
    ''' <returns>ログビルダー。</returns>
    <Extension()>
    Public Function AddZoppaLogging(builder As ILoggingBuilder,
                                    configure As Action(Of ZoppaLoggingConfiguration)) As ILoggingBuilder
        builder.AddZoppaLogging()
        builder.Services.Configure(configure)

        Return builder
    End Function

End Module