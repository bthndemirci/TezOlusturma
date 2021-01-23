using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
namespace TezOlusturma.App
{
    public class DocumentData
    {
        public List<string> Paragraflar { get; set; } = new List<string>();
        public List<Word.Paragraph> KelimeParagraflari { get; set; } = new List<Word.Paragraph>();
        public List<string> Sekiller { get; set; } = new List<string>();
        public List<string> SekilAtiflari { get; set; } = new List<string>();
        public List<string> HataliSekilAtiflari { get; set; } = new List<string>();
        public List<string> Tablolar { get; set; } = new List<string>();
        public List<string> TabloAtiflari { get; set; } = new List<string>();
        public List<string> HataliTabloAtiflari { get; set; } = new List<string>();
        public List<string> Kaynaklar { get; set; } = new List<string>();
        public List<string> KaynakAtiflari { get; set; } = new List<string>();
        public List<string> HataliKaynakAtiflari { get; set; } = new List<string>();
        public int SayfaSayisi { get; set; }
        public int ResimSayisi { get; set; }
        public int YorumSayisi { get; set; }
        public string BelgeBosluklari { get; set; }
        public string Email { get; set; }
        public string YazarBilgisi { get; set; }
        public string OlusturmaTarihi { get; set; }
        public string SonGuncelleme { get; set; }
        public string YazarAdi { get; set; }
        public string Baslik { get; set; }
        public string Aciklama { get; set; }

        public int ParagrafSayisi => Paragraflar?.Count ?? 0;
    }
}
