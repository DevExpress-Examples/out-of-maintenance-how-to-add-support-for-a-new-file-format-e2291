Imports Microsoft.VisualBasic
Imports System
Imports System.IO
#Region "#usingsEmlDocumentExporter"
Imports DevExpress.XtraRichEdit.Model
Imports DevExpress.XtraRichEdit.Internal
Imports DevExpress.XtraRichEdit.Export
Imports DevExpress.XtraRichEdit
#End Region ' #usingsEmlDocumentExporter

Namespace Eml

#Region "#EmlDocumentExporter"
    Public Class EmlDocumentExporter
        Implements IExporter(Of DocumentFormat, Boolean)
        Public Shared ReadOnly _filter As FileDialogFilter = New FileDialogFilter("E-Mail document", "eml")

#Region "IDocumentExporter Members"
        Public ReadOnly Property Filter() As FileDialogFilter _
        Implements IExporter(Of DocumentFormat, Boolean).Filter
            Get
                Return _filter
            End Get
        End Property
        Public ReadOnly Property Format() As DocumentFormat _
        Implements IExporter(Of DevExpress.XtraRichEdit.DocumentFormat, Boolean).Format
            Get
                Return EmlDocumentFormat.Id
            End Get
        End Property
        Public Function SetupSaving() As IExporterOptions _
        Implements IExporter(Of DocumentFormat, Boolean).SetupSaving
            Return New EmlDocumentExporterOptions()
        End Function

        Public Function SaveDocument(ByVal documentModel As DocumentModel, _
        ByVal stream As Stream, ByVal options As IExporterOptions) As Boolean _
        Implements IExporter(Of DevExpress.XtraRichEdit.DocumentFormat, Boolean).SaveDocument
            documentModel.InternalAPI.SaveDocumentMhtContent(stream, CType(options, EmlDocumentExporterOptions))
            Return True
        End Function
#End Region

        Public Shared Sub Register(ByVal provider As IServiceProvider)
            If provider Is Nothing Then
                Return
            End If
            Dim service As IDocumentExportManagerService = _
            TryCast(provider.GetService(GetType(IDocumentExportManagerService)), IDocumentExportManagerService)
            If Not service Is Nothing Then
                service.RegisterExporter(New EmlDocumentExporter())
            End If
        End Sub
    End Class
#End Region ' #EmlDocumentExporter

#Region "#EmlDocumentExporterOptions"
    Public Class EmlDocumentExporterOptions
        Inherits MhtDocumentExporterOptions
        Protected Overrides ReadOnly Property Format() As DocumentFormat
            Get
                Return EmlDocumentFormat.Id
            End Get
        End Property
    End Class
#End Region ' #EmlDocumentExporterOptions
End Namespace