// ============================================================================
// Project   : AnyFile Editor (TxtOrganizer)
// File      : Exceptions/Exceptions.cs
// Author    : Xavier Bento
// Version   : v1.0
// Created   : 2025-09-21
// Description: Custom exceptions used across the editor.
// ============================================================================
// File: AnyFileEditor_fixed_all/Exceptions/Exceptions.cs
// Purpose: C# implementation file in the editor application.
// Context: May interact with ThemeManager, Tabs, or file I/O
// Notes: Do not alter public API or behavior.
namespace TxtOrganizer.Exceptions
{
    /// <summary>Thrown when a file format is unsupported.</summary>
    /// <summary>FileFormatUnsupportedException — role and responsibilities within the AnyFile Editor app.</summary>
/// <remarks>Documented without behavior changes on 2025-09-21.</remarks>
    public class FileFormatUnsupportedException : System.Exception
    {
        public FileFormatUnsupportedException(string message) : base(message) {}
        public FileFormatUnsupportedException(string message, System.Exception inner) : base(message, inner) {}
    }

    /// <summary>Thrown when a document fails to load.</summary>
    /// <summary>DocumentLoadException — role and responsibilities within the AnyFile Editor app.</summary>
/// <remarks>Documented without behavior changes on 2025-09-21.</remarks>
    public class DocumentLoadException : System.Exception
    {
        public DocumentLoadException(string message) : base(message) {}
        public DocumentLoadException(string message, System.Exception inner) : base(message, inner) {}
    }

    /// <summary>Thrown when a document fails to save.</summary>
    public class DocumentSaveException : System.Exception
    {
        public DocumentSaveException(string message) : base(message) {}
        public DocumentSaveException(string message, System.Exception inner) : base(message, inner) {}
    }

    /// <summary>Thrown when a required resource (icon, image) is not found.</summary>
    public class ResourceNotFoundException : System.Exception
    {
        public ResourceNotFoundException(string message) : base(message) {}
        public ResourceNotFoundException(string message, System.Exception inner) : base(message, inner) {}
    }

    /// <summary>Thrown when a print operation cannot be completed.</summary>
    public class PrintOperationException : System.Exception
    {
        public PrintOperationException(string message) : base(message) {}
        public PrintOperationException(string message, System.Exception inner) : base(message, inner) {}
    }
}