using System;
using System.Collections.Generic;
using System.Text;

namespace DocxMould
{
    public class Replacement
    {
        public Replacement()
        {
            this._dict = new Dictionary<string, ReplaceValue>();
        }

        private IDictionary<string, ReplaceValue> _dict { get; set; }

        internal IDictionary<string, ReplaceValue> Dict
        {
            get
            {
                return this._dict;
            }
        }

        /// <summary>
        /// Add field replacement.If field already existed old value will be replaced.
        /// </summary>
        /// <param name="fieldName"></param>
        /// <param name="value"></param>
        /// <param name="style"></param>
        public void Add(string fieldName, string value, ReplaceStyle style = null)
        {
            var key = this.FieldKey(fieldName);
            var replaceValue = new ReplaceValue() { Value = value, Style = style };
            if (this._dict.ContainsKey(key))
            {
                this._dict[key] = replaceValue;
            }
            else
            {
                this._dict.Add(key, replaceValue);
            }
        }

        /// <summary>
        /// Remove field replacement.
        /// </summary>
        /// <param name="fieldName"></param>
        public void Remove(string fieldName)
        {
            var key = this.FieldKey(fieldName);
            if (this._dict.ContainsKey(key))
            {
                this._dict.Remove(key);
            }
        }

        public bool ContainsField(string fieldName)
        {
            var key = this.FieldKey(fieldName);
            return this._dict.ContainsKey(key);
        }

        private string FieldKey(string fieldName)
        {
            var key = MouldSettings.FieldPrefix + fieldName + MouldSettings.FieldSuffix;
            return key;
        }
    }

}
