Imports Microsoft.VisualBasic
Imports System
Imports System.IO
#Region "#usingsEmlDocumentImporter"
Imports DevExpress.XtraRichEdit.Model
Imports DevExpress.XtraRichEdit.Internal
Imports DevExpress.XtraRichEdit.Import
Imports DevExpress.XtraRichEdit
#End Region ' #usingsEmlDocumentImporter

Namespace Eml
#Region "#EmlDocumentFormat"
    Public Class EmlDocumentFormat
        Public Shared ReadOnly Id As DocumentFormat = New DocumentFormat(444)
    End Class
#End Region ' #EmlDocumentFormat
#Region "#EmlDocumentImporter"
    Public Class EmlDocumentImporter
        Implements IImporter(Of DocumentFormat, Boolean)
#Region "IDocumentImporter Members"
        Public ReadOnly Property Filter() As FileDialogFilter _
        Implements IImporter(Of DocumentFormat, Boolean).Filter
            Get
                Return EmlDocumentExporter._filter
            End Get
        End Property
        Public ReadOnly Property Format() As DocumentFormat _
        Implements IImporter(Of DevExpress.XtraRichEdit.DocumentFormat, Boolean).Format
            Get
                Return EmlDocumentFormat.Id
            End Get
        End Property
        Public Function SetupLoading() As IImporterOptions _
        Implements IImporter(Of DocumentFormat, Boolean).SetupLoading
            Return New EmlDocumentImporterOptions()
        End Function

        Public Function LoadDocument(ByVal documentModel As DocumentModel, ByVal stream As Stream, _
        ByVal options As IImporterOptions) As Boolean _
        Implements IImporter(Of DevExpress.XtraRichEdit.DocumentFormat, Boolean).LoadDocument
            documentModel.InternalAPI.LoadDocumentMhtContent(stream, CType(options, EmlDocumentImporterOptions))
            Return True
        End Function
#End Region

        Public Shared Sub Register(ByVal provider As IServiceProvider)
            If provider Is Nothing Then
                Return
            End If
            Dim service As IDocumentImportManagerService = _
            TryCast(provider.GetService(GetType(IDocumentImportManagerService)), IDocumentImportManagerService)
            If Not service Is Nothing Then
                service.RegisterImporter(New EmlDocumentImporter())
            End If
        End Sub
    End Class
#End Region ' #EmlDocumentImporter

#Region "#EmlDocumentImporterOptions"
    Public Class EmlDocumentImporterOptions
        Inherits MhtDocumentImporterOptions
        Protected Overrides ReadOnly Property Format() As DocumentFormat
            Get
                Return EmlDocumentFormat.Id
            End Get
        End Property
    End Class
#End Region ' #EmlDocumentImporterOptions
End Namespace