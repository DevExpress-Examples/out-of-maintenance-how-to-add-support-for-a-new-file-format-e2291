Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports DevExpress.XtraRichEdit.Export
Imports DevExpress.XtraRichEdit.Import

Namespace Eml
	Public Partial Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

			EmlDocumentExporter.Register(richEdit)
			EmlDocumentImporter.Register(richEdit)

			richEdit.LoadDocument("Hey, look at this!.eml")
		End Sub

	End Class
End Namespace