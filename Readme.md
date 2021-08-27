<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128609339/12.1.4%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/E2291)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
<!-- default file list -->
*Files to look at*:

* [DocumentExporter.cs](./CS/Export/DocumentExporter.cs) (VB: [DocumentExporter.vb](./VB/Export/DocumentExporter.vb))
* [Form1.cs](./CS/Form1.cs) (VB: [Form1.vb](./VB/Form1.vb))
* [DocumentImporter.cs](./CS/Import/DocumentImporter.cs) (VB: [DocumentImporter.vb](./VB/Import/DocumentImporter.vb))
<!-- default file list end -->
# How to add support for a new file format


<p>This example demonstrates how you can teach the RichEditControl in your application to open and save files in a custom format. Thanks to the <strong>IImporter</strong> and <strong>IExporter</strong> interfaces, this task is rather straightforward, and can be accomplished with ease.</p><p>As an example, we'll add support for a well-known format .eml. Files in this format are plain text in MIME standard, so the XtraRichEdit's built-in save and load methods for MHT content can be used. This fact simplifies the implementation, but you can write your own import/export helper classes if required.</p><p>First, provide a unique identifier for a new format. It should be a new <a href="http://documentation.devexpress.com/#Silverlight/clsDevExpressXtraRichEditDocumentFormattopic">DocumentFormat</a> enumeration member. Take a value that is large enough to exceed the used range, e.g. 444. Then, create two classes - importer and exporter, and two classes that represent options for import and export. </p><p>The class containing options for import should be derived from the <a href="http://documentation.devexpress.com/#Silverlight/clsDevExpressXtraRichEditImportDocumentImporterOptionstopic">DocumentImporterOptions</a> class or its descendants. In our case, it is the EmlDocumentImporterOptions derived from the <a href="http://documentation.devexpress.com/#Silverlight/clsDevExpressXtraRichEditImportMhtDocumentImporterOptionstopic">MhtDocumentImporterOptions</a>. </p><p>The class<strong> EmlDocumentImporter</strong> implements the <strong>IImporter<DocumentFormat, bool></strong> interface. The basic functionality is encapsulated in its LoadDocument method. To load a document in .eml format, we simply call the <strong>DocumentModel.InternalAPI.LoadDocumentMhtContent</strong> method, but for other formats you can use your custom procedures. </p><p>The class containing options for export should be derived from the <a href="http://documentation.devexpress.com/#Silverlight/clsDevExpressXtraRichEditExportDocumentExporterOptionstopic">DocumentExporterOptions</a> class or its descendants. In our case, it is the EmlDocumentExporterOptions derived from the <a href="http://documentation.devexpress.com/#Silverlight/clsDevExpressXtraRichEditExportMhtDocumentExporterOptionstopic">MhtDocumentExporterOptions</a>. </p><p>The class EmlDocumentImporter implements the <strong>IExporter<DocumentFormat, bool></strong> interface. In its SaveDocument method, the <strong>DocumentModel.InternalAPI.SaveDocumentMhtContent</strong> method is used, but you are free to use a custom procedure instead. </p><p>We have all the required objects capable of performing export-import operations. To teach the XtraRichEdit to use our objects for handling .eml files, consider the XtraRichEdit services. Loose coupling provided by the service pattern enables us to easily incorporate the desired functionality into the XtraRichEdit behavior models. The <strong>IDocumentImportManagerService</strong> has the RegisterImporter method, and the <strong>IDocumentExportManagerService</strong> has the RegisterExporter method, designed for this purpose. Register our importer and exporter classes, and the RichEditControl is capable of loading and saving files in .eml format. </p><p>Note that this example uses objects from the <strong>DevExpress.XtraRichEdit.Model</strong> and<strong> DevExpress.XtraRichEdit.Internal</strong> namespaces, not included in the documentation. The DevExpress team does not encourage the use of objects which are not listed in our documentation (help reference), and does not guarantee that these objects will be retained in future versions.</p>

<br/>


