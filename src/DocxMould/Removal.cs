using System.Collections.Generic;

namespace DocxMould
{
    public class Removal
    {
        public Removal()
        {
            this._dict = new Dictionary<string, string>();
        }

        private IDictionary<string, string> _dict { get; set; }

        internal IDictionary<string, string> Dict
        {
            get
            {
                return this._dict;
            }
        }

        /// <summary>
        /// Add section removal.If remove from already existed old remove to will be replaced.
        /// </summary>
        /// <param name="from"></param>
        /// <param name="to"></param>
        public void Add(string from, string to)
        {
            var key = this.SectionKey(from);
            var value = this.SectionKey(to);
            if (this._dict.ContainsKey(key))
            {
                this._dict[key] = value;
            }
            else
            {
                this._dict.Add(key, value);
            }
        }

        /// <summary>
        /// Remove section removal.
        /// </summary>
        /// <param name="from"></param>
        public void Remove(string from)
        {
            var key = this.SectionKey(from);
            if (this._dict.ContainsKey(key))
            {
                this._dict.Remove(key);
            }
        }

        public bool ContainsRemoval(string from)
        {
            var key = this.SectionKey(from);
            return this._dict.ContainsKey(key);
        }

        private string SectionKey(string from)
        {
            var key = MouldSettings.SectionPrefix + from + MouldSettings.SectionSuffix;
            return key;
        }
    }
}