using System;
using System.IO;

namespace DocxMould
{
    public class Mould : IDisposable
    {
        public Mould(Stream stream)
        {
            this._processor = new Processor(stream);
        }

        private bool _saved = false;

        private bool _disposed;

        private Processor _processor;

        public void ReplaceField(Replacement replacement)
        {
            if (this._saved) throw new AccessViolationException("Document has already saved can not continue processing!");
            this._processor.Replace(replacement);
        }

        public void RemoveSection(Removal removal)
        {
            if (this._saved) throw new AccessViolationException("Document has already saved can not continue processing!");
            this._processor.RemoveSection(removal);
        }

        public void Save()
        {
            if (this._saved) throw new AccessViolationException("Document has already saved can not continue processing!");
            this._processor.Save();
            this._saved = true;
        }

        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~Mould()
        {
            this.Dispose(false);
        }

        private void Dispose(bool disposing)
        {
            if (_disposed) return;
            if (disposing)
            {
                //clean up managed resource
            }
            //clean up unmanaged resource
            if (this._processor != null)
            {
                this._processor.Dispose();
                this._processor = null;
            }
            _disposed = true;
        }
    }
}
