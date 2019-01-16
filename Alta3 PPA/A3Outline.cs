using System.Collections.Generic;
using System.IO;
using YamlDotNet.Serialization;

namespace Alta3_PPA {
    public class A3Outline {
        #region Outline Properites
        public enum Metadata {
            NAME,
            FILENAME,
            HASLABS,
            HASSLIDES,
            HASVIDEOS,
            WEBURL
        }


        public string Course { get; set; }
        public string Filename { get; set; }
        public bool HasLabs { get; set; }
        public bool HasSlides { get; set; }
        public bool HasVideos { get; set; }
        public string Weburl { get; set; }
        public List<A3Chapter> Chapters { get; set; }
        #endregion

        public A3Outline()
        {
            Course = null;
            Filename = null;
            HasLabs = false;
            HasSlides = false;
            HasVideos = false;
            Weburl = null;
            Chapters = new List<A3Chapter>();
        }

    }
}
