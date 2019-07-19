using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace DocxMould
{
    internal class Processor : IDisposable
    {
        public Processor(Stream stream)
        {
            this._stream = stream;
            this._docx = WordprocessingDocument.Open(stream, true);
        }

        private bool _disposed;

        public Stream _stream;

        public WordprocessingDocument _docx;

        public void Replace(Replacement replacement)
        {
            throw new NotImplementedException();
        }

        public void RemoveSection(Removal removal)
        {
            throw new NotImplementedException();
        }

        public void Save()
        {
            this.RemoveAllTags();
            throw new NotImplementedException();
        }

        private void RemoveAllTags()
        {
            throw new NotImplementedException();
        }

        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~Processor()
        {
            this.Dispose(false);
        }

        private void Dispose(bool disposing)
        {
            if (_disposed) return;
            if (disposing)
            {
                //clean up managed resource
                this._stream = null;
            }
            //clean up unmanaged resource
            if (this._docx != null)
            {
                this._docx.Close();
                this._docx.Dispose();
            }
            _disposed = true;
        }
    }
}
