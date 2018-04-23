Imports Microsoft.VisualBasic
Imports System
Imports System.IO
#Region "#usingsEmlDocumentImporter"
Imports DevExpress.XtraRichEdit.Model
Imports DevExpress.XtraRichEdit.Internal
Imports DevExpress.XtraRichEdit.Import
Imports DevExpress.XtraRichEdit
Imports DevExpress.Office.Internal
Imports DevExpress.Office.Import
#End Region ' #usingsEmlDocumentImporter

Namespace Eml
#Region "#EmlDocumentFormat"
    Public NotInheritable Class EmlDocumentFormat
        Public Shared ReadOnly Id As New DocumentFormat(444)
    End Class
#End Region ' #EmlDocumentFormat
#Region "#EmlDocumentImporter"
    Public Class EmlDocumentImporter
        Implements IImporter(Of DocumentFormat, Boolean)

#Region "IDocumentImporter Members"
        Public ReadOnly Property Filter() As FileDialogFilter Implements IImporter(Of DocumentFormat, Boolean).Filter
            Get
                Return EmlDocumentExporter.filter_Renamed
            End Get
        End Property
        Public ReadOnly Property Format() As DocumentFormat _
        Implements IImporter(Of DevExpress.XtraRichEdit.DocumentFormat, Boolean).Format
            Get
                Return EmlDocumentFormat.Id
            End Get
        End Property
        Public Function SetupLoading() As IImporterOptions Implements IImporter(Of DocumentFormat, Boolean).SetupLoading
            Return New EmlDocumentImporterOptions()
        End Function

        Public Function LoadDocument(ByVal documentModel As DevExpress.Office.IDocumentModel, ByVal stream As Stream, ByVal options As IImporterOptions) As Boolean Implements IImporter(Of DevExpress.XtraRichEdit.DocumentFormat, Boolean).LoadDocument
            Return LoadDocument(documentModel, stream, options, True)
        End Function

        Public Function LoadDocument(ByVal documentModel As DevExpress.Office.IDocumentModel, ByVal stream As Stream, ByVal options As IImporterOptions, ByVal leaveOpen As Boolean) As Boolean Implements IImporter(Of DevExpress.XtraRichEdit.DocumentFormat, Boolean).LoadDocument
            Dim model As DocumentModel = CType(documentModel, DocumentModel)
            model.InternalAPI.LoadDocumentMhtContent(stream, CType(options, EmlDocumentImporterOptions))
            Return True
        End Function

#End Region

        Public Shared Sub Register(ByVal provider As IServiceProvider)
            If provider Is Nothing Then
                Return
            End If
            Dim service As IDocumentImportManagerService = TryCast(provider.GetService(GetType(IDocumentImportManagerService)), IDocumentImportManagerService)
            If service IsNot Nothing Then
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