using System;
using System.IO;
#region #usingsEmlDocumentImporter
using DevExpress.XtraRichEdit.Model;
using DevExpress.XtraRichEdit.Internal;
using DevExpress.XtraRichEdit.Import;
using DevExpress.XtraRichEdit;
using DevExpress.Office.Internal;
using DevExpress.Office.Import;
#endregion #usingsEmlDocumentImporter

namespace Eml {
    #region #EmlDocumentFormat
    public static class EmlDocumentFormat {
        public static readonly DocumentFormat Id = new DocumentFormat(444);
    }
    #endregion #EmlDocumentFormat
    #region #EmlDocumentImporter
    public class EmlDocumentImporter : IImporter<DocumentFormat, bool> {
        #region IDocumentImporter Members
        public FileDialogFilter Filter { get { return EmlDocumentExporter.filter; } }
        public DocumentFormat Format { get { return EmlDocumentFormat.Id; } }
        public IImporterOptions SetupLoading() {
            return new EmlDocumentImporterOptions();
        }

        public bool LoadDocument(DevExpress.Office.IDocumentModel documentModel, Stream stream, IImporterOptions options) {
            return LoadDocument(documentModel, stream, options, true);

        }

        public bool LoadDocument(DevExpress.Office.IDocumentModel documentModel, Stream stream, IImporterOptions options, bool leaveOpen) {
            DocumentModel model = (DocumentModel)documentModel;
            model.InternalAPI.LoadDocumentMhtContent(stream, (EmlDocumentImporterOptions)options);
            return true;
        }

        #endregion

        public static void Register(IServiceProvider provider) {
            if(provider == null)
                return;
            IDocumentImportManagerService service = provider.GetService(typeof(IDocumentImportManagerService)) as IDocumentImportManagerService;
            if(service != null)
                service.RegisterImporter(new EmlDocumentImporter());
        }

    }
    #endregion #EmlDocumentImporter

    #region #EmlDocumentImporterOptions
    public class EmlDocumentImporterOptions : MhtDocumentImporterOptions {
        protected override DocumentFormat Format { get { return EmlDocumentFormat.Id; } }
    }
    #endregion #EmlDocumentImporterOptions
}