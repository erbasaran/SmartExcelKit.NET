using System.Runtime.Serialization;

namespace SmartExcelKit.Exceptions;

/// <summary>
/// Base exception class for all errors that occur within the SmartExcelKit library.
/// </summary>
[Serializable]
public class SmartExcelException : Exception
{
    /// <summary>
    /// Gets the unique error code associated with this exception.
    /// </summary>
    public string? ErrorCode { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="SmartExcelException"/> class.
    /// </summary>
    public SmartExcelException() : base() { }

    /// <summary>
    /// Initializes a new instance of the <see cref="SmartExcelException"/> class with a specified error message.
    /// </summary>
    /// <param name="message">The message that describes the error.</param>
    public SmartExcelException(string message) : base(message) { }

    /// <summary>
    /// Initializes a new instance of the <see cref="SmartExcelException"/> class with a specified error message and a reference to the inner exception that is the cause of this exception.
    /// </summary>
    /// <param name="message">The message that describes the error.</param>
    /// <param name="innerException">The exception that is the cause of the current exception.</param>
    public SmartExcelException(string message, Exception innerException) : base(message, innerException) { }

    /// <summary>
    /// Initializes a new instance of the <see cref="SmartExcelException"/> class with a specified error message and error code.
    /// </summary>
    /// <param name="message">The message that describes the error.</param>
    /// <param name="errorCode">The application-specific error code.</param>
    public SmartExcelException(string message, string errorCode) : base(message)
    {
        ErrorCode = errorCode;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="SmartExcelException"/> class with a specified error message, error code, and inner exception.
    /// </summary>
    /// <param name="message">The message that describes the error.</param>
    /// <param name="errorCode">The application-specific error code.</param>
    /// <param name="innerException">The inner exception that is the cause of the current exception.</param>
    public SmartExcelException(string message, string errorCode, Exception innerException) : base(message, innerException)
    {
        ErrorCode = errorCode;
    }

    /// <summary>
    /// Serialization constructor.
    /// </summary>
    /// <param name="info">The serialization info.</param>
    /// <param name="context">The streaming context.</param>
    protected SmartExcelException(SerializationInfo info, StreamingContext context) : base(info, context)
    {
        ErrorCode = info.GetString(nameof(ErrorCode));
    }

    /// <inheritdoc />
    public override void GetObjectData(SerializationInfo info, StreamingContext context)
    {
        if (info == null) throw new ArgumentNullException(nameof(info));
        base.GetObjectData(info, context);
        info.AddValue(nameof(ErrorCode), ErrorCode);
    }
}

/// <summary>
/// Exception thrown for workbook-related errors.
/// </summary>
[Serializable]
public class WorkbookException : SmartExcelException
{
    /// <inheritdoc />
    public WorkbookException() : base() { }
    /// <inheritdoc />
    public WorkbookException(string message) : base(message) { }
    /// <inheritdoc />
    public WorkbookException(string message, Exception innerException) : base(message, innerException) { }
    /// <inheritdoc />
    public WorkbookException(string message, string errorCode) : base(message, errorCode) { }
    /// <inheritdoc />
    public WorkbookException(string message, string errorCode, Exception innerException) : base(message, errorCode, innerException) { }
    /// <inheritdoc />
    protected WorkbookException(SerializationInfo info, StreamingContext context) : base(info, context) { }
}

/// <summary>
/// Exception thrown for worksheet-related errors.
/// </summary>
[Serializable]
public class WorksheetException : SmartExcelException
{
    /// <inheritdoc />
    public WorksheetException() : base() { }
    /// <inheritdoc />
    public WorksheetException(string message) : base(message) { }
    /// <inheritdoc />
    public WorksheetException(string message, Exception innerException) : base(message, innerException) { }
    /// <inheritdoc />
    public WorksheetException(string message, string errorCode) : base(message, errorCode) { }
    /// <inheritdoc />
    public WorksheetException(string message, string errorCode, Exception innerException) : base(message, errorCode, innerException) { }
    /// <inheritdoc />
    protected WorksheetException(SerializationInfo info, StreamingContext context) : base(info, context) { }
}

/// <summary>
/// Exception thrown for cell-related errors.
/// </summary>
[Serializable]
public class CellException : SmartExcelException
{
    /// <inheritdoc />
    public CellException() : base() { }
    /// <inheritdoc />
    public CellException(string message) : base(message) { }
    /// <inheritdoc />
    public CellException(string message, Exception innerException) : base(message, innerException) { }
    /// <inheritdoc />
    public CellException(string message, string errorCode) : base(message, errorCode) { }
    /// <inheritdoc />
    public CellException(string message, string errorCode, Exception innerException) : base(message, errorCode, innerException) { }
    /// <inheritdoc />
    protected CellException(SerializationInfo info, StreamingContext context) : base(info, context) { }
}

/// <summary>
/// Exception thrown for formula parser/evaluation errors.
/// </summary>
[Serializable]
public class FormulaException : SmartExcelException
{
    /// <inheritdoc />
    public FormulaException() : base() { }
    /// <inheritdoc />
    public FormulaException(string message) : base(message) { }
    /// <inheritdoc />
    public FormulaException(string message, Exception innerException) : base(message, innerException) { }
    /// <inheritdoc />
    public FormulaException(string message, string errorCode) : base(message, errorCode) { }
    /// <inheritdoc />
    public FormulaException(string message, string errorCode, Exception innerException) : base(message, errorCode, innerException) { }
    /// <inheritdoc />
    protected FormulaException(SerializationInfo info, StreamingContext context) : base(info, context) { }
}

/// <summary>
/// Exception thrown for import errors.
/// </summary>
[Serializable]
public class ImportException : SmartExcelException
{
    /// <inheritdoc />
    public ImportException() : base() { }
    /// <inheritdoc />
    public ImportException(string message) : base(message) { }
    /// <inheritdoc />
    public ImportException(string message, Exception innerException) : base(message, innerException) { }
    /// <inheritdoc />
    public ImportException(string message, string errorCode) : base(message, errorCode) { }
    /// <inheritdoc />
    public ImportException(string message, string errorCode, Exception innerException) : base(message, errorCode, innerException) { }
    /// <inheritdoc />
    protected ImportException(SerializationInfo info, StreamingContext context) : base(info, context) { }
}

/// <summary>
/// Exception thrown for export errors.
/// </summary>
[Serializable]
public class ExportException : SmartExcelException
{
    /// <inheritdoc />
    public ExportException() : base() { }
    /// <inheritdoc />
    public ExportException(string message) : base(message) { }
    /// <inheritdoc />
    public ExportException(string message, Exception innerException) : base(message, innerException) { }
    /// <inheritdoc />
    public ExportException(string message, string errorCode) : base(message, errorCode) { }
    /// <inheritdoc />
    public ExportException(string message, string errorCode, Exception innerException) : base(message, errorCode, innerException) { }
    /// <inheritdoc />
    protected ExportException(SerializationInfo info, StreamingContext context) : base(info, context) { }
}

/// <summary>
/// Exception thrown for validation-related errors.
/// </summary>
[Serializable]
public class ValidationException : SmartExcelException
{
    /// <inheritdoc />
    public ValidationException() : base() { }
    /// <inheritdoc />
    public ValidationException(string message) : base(message) { }
    /// <inheritdoc />
    public ValidationException(string message, Exception innerException) : base(message, innerException) { }
    /// <inheritdoc />
    public ValidationException(string message, string errorCode) : base(message, errorCode) { }
    /// <inheritdoc />
    public ValidationException(string message, string errorCode, Exception innerException) : base(message, errorCode, innerException) { }
    /// <inheritdoc />
    protected ValidationException(SerializationInfo info, StreamingContext context) : base(info, context) { }
}

/// <summary>
/// Exception thrown for serialization-related errors.
/// </summary>
[Serializable]
public class SerializationException : SmartExcelException
{
    /// <inheritdoc />
    public SerializationException() : base() { }
    /// <inheritdoc />
    public SerializationException(string message) : base(message) { }
    /// <inheritdoc />
    public SerializationException(string message, Exception innerException) : base(message, innerException) { }
    /// <inheritdoc />
    public SerializationException(string message, string errorCode) : base(message, errorCode) { }
    /// <inheritdoc />
    public SerializationException(string message, string errorCode, Exception innerException) : base(message, errorCode, innerException) { }
    /// <inheritdoc />
    protected SerializationException(SerializationInfo info, StreamingContext context) : base(info, context) { }
}

/// <summary>
/// Exception thrown for encryption-related errors.
/// </summary>
[Serializable]
public class EncryptionException : SmartExcelException
{
    /// <inheritdoc />
    public EncryptionException() : base() { }
    /// <inheritdoc />
    public EncryptionException(string message) : base(message) { }
    /// <inheritdoc />
    public EncryptionException(string message, Exception innerException) : base(message, innerException) { }
    /// <inheritdoc />
    public EncryptionException(string message, string errorCode) : base(message, errorCode) { }
    /// <inheritdoc />
    public EncryptionException(string message, string errorCode, Exception innerException) : base(message, errorCode, innerException) { }
    /// <inheritdoc />
    protected EncryptionException(SerializationInfo info, StreamingContext context) : base(info, context) { }
}

/// <summary>
/// Exception thrown for protection-related errors.
/// </summary>
[Serializable]
public class ProtectionException : SmartExcelException
{
    /// <inheritdoc />
    public ProtectionException() : base() { }
    /// <inheritdoc />
    public ProtectionException(string message) : base(message) { }
    /// <inheritdoc />
    public ProtectionException(string message, Exception innerException) : base(message, innerException) { }
    /// <inheritdoc />
    public ProtectionException(string message, string errorCode) : base(message, errorCode) { }
    /// <inheritdoc />
    public ProtectionException(string message, string errorCode, Exception innerException) : base(message, errorCode, innerException) { }
    /// <inheritdoc />
    protected ProtectionException(SerializationInfo info, StreamingContext context) : base(info, context) { }
}

/// <summary>
/// Exception thrown for chart-related errors.
/// </summary>
[Serializable]
public class ChartException : SmartExcelException
{
    /// <inheritdoc />
    public ChartException() : base() { }
    /// <inheritdoc />
    public ChartException(string message) : base(message) { }
    /// <inheritdoc />
    public ChartException(string message, Exception innerException) : base(message, innerException) { }
    /// <inheritdoc />
    public ChartException(string message, string errorCode) : base(message, errorCode) { }
    /// <inheritdoc />
    public ChartException(string message, string errorCode, Exception innerException) : base(message, errorCode, innerException) { }
    /// <inheritdoc />
    protected ChartException(SerializationInfo info, StreamingContext context) : base(info, context) { }
}

/// <summary>
/// Exception thrown for image-related errors.
/// </summary>
[Serializable]
public class ImageException : SmartExcelException
{
    /// <inheritdoc />
    public ImageException() : base() { }
    /// <inheritdoc />
    public ImageException(string message) : base(message) { }
    /// <inheritdoc />
    public ImageException(string message, Exception innerException) : base(message, innerException) { }
    /// <inheritdoc />
    public ImageException(string message, string errorCode) : base(message, errorCode) { }
    /// <inheritdoc />
    public ImageException(string message, string errorCode, Exception innerException) : base(message, errorCode, innerException) { }
    /// <inheritdoc />
    protected ImageException(SerializationInfo info, StreamingContext context) : base(info, context) { }
}

/// <summary>
/// Exception thrown for document metadata-related errors.
/// </summary>
[Serializable]
public class MetadataException : SmartExcelException
{
    /// <inheritdoc />
    public MetadataException() : base() { }
    /// <inheritdoc />
    public MetadataException(string message) : base(message) { }
    /// <inheritdoc />
    public MetadataException(string message, Exception innerException) : base(message, innerException) { }
    /// <inheritdoc />
    public MetadataException(string message, string errorCode) : base(message, errorCode) { }
    /// <inheritdoc />
    public MetadataException(string message, string errorCode, Exception innerException) : base(message, errorCode, innerException) { }
    /// <inheritdoc />
    protected MetadataException(SerializationInfo info, StreamingContext context) : base(info, context) { }
}

/// <summary>
/// Exception thrown for parsing-related errors.
/// </summary>
[Serializable]
public class ParsingException : SmartExcelException
{
    /// <inheritdoc />
    public ParsingException() : base() { }
    /// <inheritdoc />
    public ParsingException(string message) : base(message) { }
    /// <inheritdoc />
    public ParsingException(string message, Exception innerException) : base(message, innerException) { }
    /// <inheritdoc />
    public ParsingException(string message, string errorCode) : base(message, errorCode) { }
    /// <inheritdoc />
    public ParsingException(string message, string errorCode, Exception innerException) : base(message, errorCode, innerException) { }
    /// <inheritdoc />
    protected ParsingException(SerializationInfo info, StreamingContext context) : base(info, context) { }
}
