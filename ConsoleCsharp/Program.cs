using Microsoft.Extensions.Logging;
using ZoppaLoggingExtensions;

using var loggerFactory = LoggerFactory.Create(builder => {
    builder.AddZoppaLogging(
        (config) => config.MinimumLogLevel = LogLevel.Trace
    );
    builder.SetMinimumLevel(LogLevel.Trace);
});

ILogger logger = loggerFactory.CreateLogger<Program>();
using (logger.BeginScope("1")) {
    logger.LogDebug(1, "Does this line get hit?");
    using (logger.BeginScope("2")) {
        logger.LogInformation(3, "Nothing to see here.");
        logger.LogWarning(5, "Warning... that was odd.");
    }
    logger.LogError(7, "Oops, there was an error.");
}
logger.LogTrace(5, "== 120.");