using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
            var runs = this._docx.MainDocumentPart.Document.Body.Descendants<Run>().ToList();
            var dict = replacement.Dict;
            var keys = dict.Keys.ToList();
            for (int i = 0; i < keys.Count; i++)
            {
                var replaceKey = keys[i];
                var replaceValue = dict[replaceKey];
                for (int j = 0; j < runs.Count; j++)
                {
                    var run = runs[j];
                    var text = run.InnerText;
                    if (text.Contains(replaceKey))
                    {
                        var newText = text.Replace(replaceKey, replaceValue.Value);
                        run.Elements<Text>().First().Text = newText;
                        this.SetCustomerStyle(run, replaceValue.Style);
                    }
                }
            }
        }

        private void SetCustomerStyle(Run run, ReplaceStyle style)
        {
            if (style == null) return;
            if (string.IsNullOrEmpty(style.FontFamily)) run.RunProperties.RunFonts = new RunFonts() { Ascii = style.FontFamily };
            if (style.FontSize != null) run.RunProperties.FontSize = new FontSize() { Val = new StringValue(style.FontSize.ToString()) };
            if (style.Bold != null) run.RunProperties.Bold = new Bold() { Val = new OnOffValue(style.Bold.Value) };
            if (style.Capitalized != null) run.RunProperties.Caps = new Caps() { Val = new OnOffValue(style.Capitalized.Value) };
            if (style.DoubleStrikeThrough != null) run.RunProperties.DoubleStrike = new DoubleStrike() { Val = new OnOffValue(style.DoubleStrikeThrough.Value) };
            if (style.Embossed != null) run.RunProperties.Emboss = new Emboss() { Val = new OnOffValue(style.Embossed.Value) };
            if (style.Imprinted != null) run.RunProperties.Imprint = new Imprint() { Val = new OnOffValue(style.Imprinted.Value) };
            if (style.Italic != null) run.RunProperties.Italic = new Italic() { Val = new OnOffValue(style.Italic.Value) };
            if (style.Shadowed != null) run.RunProperties.Shadow = new Shadow() { Val = new OnOffValue(style.Shadowed.Value) };
            if (style.SmallCaps != null) run.RunProperties.SmallCaps = new SmallCaps() { Val = new OnOffValue(style.SmallCaps.Value) };
            if (style.StrikeThrough != null) run.RunProperties.Strike = new Strike() { Val = new OnOffValue(style.StrikeThrough.Value) };
            if (style.VerticalPosition != null) run.RunProperties.VerticalTextAlignment = new VerticalTextAlignment() { Val = (VerticalPositionValues)style.VerticalPosition.Value };
        }

        public void RemoveSection(Removal removal)
        {
            var dict = removal.Dict;
            if (dict.Keys.Count != 0)
            {
                var keys = dict.Keys.ToList();
                for (int i = 0; i < keys.Count; i++)
                {
                    var from = keys[i];
                    var to = dict[from];
                    this.ReomveElementsBySection(from, to);
                }
            }
        }

        private void ReomveElementsBySection(string from, string to)
        {
            var body = this._docx.MainDocumentPart.Document.Body;
            while (body.InnerText.Contains(from))
            {
                if (!body.InnerText.Contains(to))
                {
                    break;
                }
                var mark = false;
                var elements = this._docx.MainDocumentPart.Document.Body.Elements().ToList();
                for (int i = 0; i < elements.Count; i++)
                {
                    var elem = elements[i];
                    if (elem is Paragraph && elem.InnerText.Contains(from)) mark = true;
                    if (mark) elem.Remove();
                    if (mark && elem is Paragraph && elem.InnerText.Contains(to)) break;
                }
            }
        }

        public void Save()
        {
            this.RemoveAllSectionTags();
            this._docx.MainDocumentPart.Document.Save();
        }

        private void RemoveAllSectionTags()
        {
            var prefix = MouldSettings.SectionPrefix;
            var suffix = MouldSettings.SectionSuffix;
            var paragraphs = this._docx.MainDocumentPart.Document.Body.Elements<Paragraph>().ToList();
            for (int i = 0; i < paragraphs.Count; i++)
            {
                var para = paragraphs[i];
                if (string.IsNullOrEmpty(para.InnerText)) continue;
                if (para.InnerText.StartsWith(prefix) && para.InnerText.EndsWith(suffix))
                {
                    para.Remove();
                }
            }
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
