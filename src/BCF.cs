using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;
using System.IO.Compression;
using System.Drawing;
using System.Drawing.Imaging;

namespace BCFReader.Note
{
    class BCFNote
    {
        public BCFNote(DirectoryInfo dI)
        {
            string markupPath = dI.EnumerateFiles("markup.bcf").First().FullName;
            _markup = ReadMarkup(markupPath);

            if (dI.EnumerateFiles("snapshot.png").Count() != 0)
            {
                string picturePath = dI.EnumerateFiles("snapshot.png").First().FullName;
                _picture = CreateNonIndexedImage(picturePath);
            }
            else
            {
                _picture = null;
            }

        }

        private Image CreateNonIndexedImage(string path)
        {
            Bitmap targetImage = null;

            using (var sourceImage = Image.FromFile(path))
            {
                targetImage = new Bitmap(sourceImage.Width, sourceImage.Height, PixelFormat.Format32bppArgb);
                using (Graphics canvas = Graphics.FromImage(targetImage))
                {
                    canvas.DrawImageUnscaled(sourceImage, 0, 0);
                }

            }

            return targetImage;
        }

        private Image _picture;
        public Image Picture
        {
            get { return _picture; }
        }

        private Markup _markup;
        public Markup Markup
        {
            get { return _markup; }
            //set { _markup = value; }
        }

        private Note.Markup ReadMarkup(string filePath)
        {
            Note.Markup markup = new Note.Markup();

            using (Stream objStream = new FileStream(filePath, FileMode.Open))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(Note.Markup));

                StreamReader sr = new StreamReader(objStream);

                XmlReaderSettings settings = new XmlReaderSettings();
                settings.IgnoreComments = true;


                using (XmlReader reader = XmlReader.Create(sr, settings))
                {
                    markup = (Note.Markup)serializer.Deserialize(reader);
                }
            }

            return markup;
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Xml", "4.0.30319.34234")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "", IsNullable = true)]
    public partial class Header
    {
        private HeaderFile[] fileField;
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("File", Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public HeaderFile[] File
        {
            get
            {
                return this.fileField;
            }
            set
            {
                this.fileField = value;
            }
        }
    }
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Xml", "4.0.30319.34234")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class HeaderFile
    {
        private string filenameField;
        private System.DateTime dateField;
        private bool dateFieldSpecified;
        private string ifcProjectField;
        private string ifcSpatialStructureElementField;
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string Filename
        {
            get
            {
                return this.filenameField;
            }
            set
            {
                this.filenameField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public System.DateTime Date
        {
            get
            {
                return this.dateField;
            }
            set
            {
                this.dateField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool DateSpecified
        {
            get
            {
                return this.dateFieldSpecified;
            }
            set
            {
                this.dateFieldSpecified = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string IfcProject
        {
            get
            {
                return this.ifcProjectField;
            }
            set
            {
                this.ifcProjectField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string IfcSpatialStructureElement
        {
            get
            {
                return this.ifcSpatialStructureElementField;
            }
            set
            {
                this.ifcSpatialStructureElementField = value;
            }
        }
    }
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Xml", "4.0.30319.34234")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "", IsNullable = true)]
    public partial class Topic
    {
        private string referenceLinkField;
        private string titleField;
        private string guidField;
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string ReferenceLink
        {
            get
            {
                return this.referenceLinkField;
            }
            set
            {
                this.referenceLinkField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string Title
        {
            get
            {
                return this.titleField;
            }
            set
            {
                this.titleField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string Guid
        {
            get
            {
                return this.guidField;
            }
            set
            {
                this.guidField = value;
            }
        }
    }
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Xml", "4.0.30319.34234")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "", IsNullable = true)]
    public partial class Comment
    {
        private string verbalStatusField;
        private CommentStatus statusField;
        private System.DateTime dateField;
        private string authorField;
        private string comment1Field;
        private CommentTopic topicField;
        private string guidField;
        public Comment()
        {
            this.statusField = CommentStatus.Unknown;
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string VerbalStatus
        {
            get
            {
                return this.verbalStatusField;
            }
            set
            {
                this.verbalStatusField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public CommentStatus Status
        {
            get
            {
                return this.statusField;
            }
            set
            {
                this.statusField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public System.DateTime Date
        {
            get
            {
                return this.dateField;
            }
            set
            {
                this.dateField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string Author
        {
            get
            {
                return this.authorField;
            }
            set
            {
                this.authorField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("Comment", Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string Comment1
        {
            get
            {
                return this.comment1Field;
            }
            set
            {
                this.comment1Field = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public CommentTopic Topic
        {
            get
            {
                return this.topicField;
            }
            set
            {
                this.topicField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string Guid
        {
            get
            {
                return this.guidField;
            }
            set
            {
                this.guidField = value;
            }
        }
    }
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Xml", "4.0.30319.34234")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "", IsNullable = false)]
    public enum CommentStatus
    {
        /// <remarks/>
        Error,
        /// <remarks/>
        Warning,
        /// <remarks/>
        Info,
        /// <remarks/>
        Unknown,
    }
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Xml", "4.0.30319.34234")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class CommentTopic
    {
        private string guidField;
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string Guid
        {
            get
            {
                return this.guidField;
            }
            set
            {
                this.guidField = value;
            }
        }
    }
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Xml", "4.0.30319.34234")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "", IsNullable = false)]
    public partial class Markup
    {
        private HeaderFile[] headerField;
        private Topic topicField;
        private Comment[] commentField;
        /// <remarks/>
        [System.Xml.Serialization.XmlArrayAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        [System.Xml.Serialization.XmlArrayItemAttribute("File", Form = System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable = false)]
        public HeaderFile[] Header
        {
            get
            {
                return this.headerField;
            }
            set
            {
                this.headerField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public Topic Topic
        {
            get
            {
                return this.topicField;
            }
            set
            {
                this.topicField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("Comment", Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public Comment[] Comment
        {
            get
            {
                return this.commentField;
            }
            set
            {
                this.commentField = value;
            }
        }
    }
}
