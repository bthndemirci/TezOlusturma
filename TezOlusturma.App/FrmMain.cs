using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DevExpress.XtraBars.Ribbon;
using Task = System.Threading.Tasks.Task;
using Word = Microsoft.Office.Interop.Word;

namespace TezOlusturma.App
{
    public partial class FrmMain : RibbonForm
    {
        private Word.Application _wordApplication;
        private Word.Document _wordDocument;
        private DocumentData _documentData;

        public FrmMain()
        {
            InitializeComponent();
            btnOpenFile.ItemClick += OpenFile;

        }
        private void OpenFile(object sender, ItemClickEventArgs e)
        {

            comFileDialog.InitialDirectory = "c:\\";
            comFileDialog.Title = @"Bir Word Dosyası Seçin";
            comFileDialog.Filter = @"Word Dosyası (*.doc)|*.docx";
            comFileDialog.FilterIndex = 2;
            comFileDialog.RestoreDirectory = true;
            if (comFileDialog.ShowDialog() != DialogResult.OK) return;

            DxWaitForm.Show(this);
            Text = comFileDialog.SafeFileName + @" | Tez Doğrulama Yazılımı";
            txtDocument.LoadDocument(comFileDialog.FileName);
            _wordApplication = new Word.Application();
            _wordDocument = _wordApplication.Documents.Open(comFileDialog.FileName);
            _documentData = new DocumentData();
            _ = AnalyzeDocument();
            LoadForm();
            DxWaitForm.Close();
        }

        private void LoadForm()
        {
            _documentData.OlusturmaTarihi = txtDocument.Document.DocumentProperties.Created.ToString("dd.MM.yyyy");
            _documentData.SonGuncelleme = txtDocument.Document.DocumentProperties.Modified.ToString("dd.MM.yyyy");
            _documentData.YazarBilgisi = txtDocument.Document.DocumentProperties.Creator;
            _documentData.Baslik = txtDocument.Document.DocumentProperties.Title;
            _documentData.Aciklama = txtDocument.Document.DocumentProperties.Description;

            var result = new List<Result>
            {
                
                new Result
                {
                    Name = "Yazar",
                    Value = _documentData.YazarAdi
                },
                new Result
                {
                    Name = "Email",
                    Value = _documentData.Email
                },
                new Result
                {
                    Name = "Yazar Bilgisi",
                    Value = _documentData.YazarBilgisi
                },
                new Result
                {
                    Name = "Başlık",
                    Value = _documentData.Baslik
                },
                new Result
                {
                    Name = "Açıklama",
                    Value = _documentData.Aciklama
                },

                new Result
                {
                    Name = "Sayfa Sayısı",
                    Value = _documentData.SayfaSayisi.ToString()
                },
                new Result
                {
                    Name = "Paragraf Sayısı",
                    Value = _documentData.ParagrafSayisi.ToString()
                },
                new Result
                {
                    Name = "Yorum Sayısı",
                    Value = _documentData.YorumSayisi.ToString()
                },
                new Result
                {
                    Name = "Resim Sayısı",
                    Value = _documentData.ResimSayisi.ToString()
                },
                new Result
                {
                    Name = "Belge Boşluklari",
                    Value = _documentData.BelgeBosluklari
                },
                new Result
                {
                    Name = "Oluşturma Tarihi",
                    Value = _documentData.OlusturmaTarihi
                },
                new Result
                {
                    Name = "Son Güncelleme Tarihi",
                    Value = _documentData.SonGuncelleme
                }
            };
            gridInfo.DataSource = result;
            lblKaynakSayisi.Text = _documentData.Kaynaklar.Count.ToString("#,##0");
            lblKaynakAtif.Text = _documentData.KaynakAtiflari.Count.ToString("#,##0");

            lblTabloSayisi.Text = _documentData.Tablolar.Count.ToString("#,##0");
            lblTabloAtif.Text = _documentData.TabloAtiflari.Count.ToString("#,##0");

            lblSekilSayisi.Text = _documentData.Sekiller.Count.ToString("#,##0");
            lblSekilAtif.Text = _documentData.SekilAtiflari.Count.ToString("#,##0");

            gridKaynak.DataSource = _documentData.Kaynaklar.Select(x => new
            {
                Kaynak = x,
                KaynakAtif = _documentData.KaynakAtiflari.Contains(x) ? "Atıf Var" : "Atıf Yok"
            });
            gridTablo.DataSource = _documentData.Tablolar.Select(x => new
            {
                Tablo = x,
                TabloAtif = _documentData.TabloAtiflari.Contains(x) ? "Atıf Var" : "Atıf Yok"
            });
            gridSekil.DataSource = _documentData.Sekiller.Select(x => new
            {
                Sekil = x,
                SekilAtif = _documentData.SekilAtiflari.Contains(x) ? "Atıf Var" : "Atıf Yok"
            });

        }

        private async Task AnalyzeDocument()
        {
            _documentData.SayfaSayisi = txtDocument.DocumentLayout.GetPageCount();
            _documentData.ResimSayisi = txtDocument.Document.Sections.Count;
            _documentData.YorumSayisi = txtDocument.Document.Comments.Count;
            _documentData.Email = txtDocument.Document.DocumentProperties.Creator;
            var left = (txtDocument.Document.Sections[0].Margins.Left / 10).ToString("0.00"); ;
            var right = (txtDocument.Document.Sections[0].Margins.Right / 10).ToString("0.00"); ;
            var bottom = (txtDocument.Document.Sections[0].Margins.Bottom / 10).ToString("0.00"); ;
            var top = (txtDocument.Document.Sections[0].Margins.Top / 10).ToString("0.00"); ;
            var paddings = $"(MM) Üst : {top}| Sağ : {right}| Alt : {bottom}| Sol : {left}";
            _documentData.BelgeBosluklari = paddings;

            _documentData.OlusturmaTarihi = txtDocument.Document.DocumentProperties.Created.ToString("dd.MM.yyyy");
            _documentData.SonGuncelleme = txtDocument.Document.DocumentProperties.Modified.ToString("dd.MM.yyyy");
            _documentData.YazarBilgisi = txtDocument.Document.DocumentProperties.Creator;
            _documentData.Baslik = txtDocument.Document.DocumentProperties.Title;
            _documentData.Aciklama= txtDocument.Document.DocumentProperties.Description;

            var kaynaklar = Kaynaklar();
            var tablolar = Tablolar();
            var sekiller = Sekiller();
            var yazar = Yazar();
            await Task.WhenAll(yazar,kaynaklar, tablolar, sekiller);
        }


        string newRow = "\r";
        string newLine = "\n";

        #region Yazar
        private async Task Yazar()
        {

            var author = "";
            var splitted = txtDocument.Document.Text.Split(new[] { "Yazarı: " }, StringSplitOptions.None);
            if (splitted.Length <= 1)
            {
                author = "'Yazarı:' ibaresi eklenmemiş";
            }
            else
            {
                splitted = splitted[1].Split(new[] { "\r\n" }, StringSplitOptions.None);
                foreach (var value in splitted)
                {
                    if (value.TrimStart().TrimEnd() == "") continue;

                    author = value.TrimStart().TrimEnd();
                    break;
                }
            }
            _documentData.YazarAdi = author;
            await Task.Yield();
        }
        #endregion

        #region Kaynaklar ve Kaynaklara Atıflar
        private async Task Kaynaklar()
        {
            #region Kaynaklar
            foreach (Word.Paragraph paragraph in _wordDocument.Paragraphs)
            {
                var text = paragraph.Range.Text;
                text = text.TrimStart().TrimEnd().Replace(newLine, "").Replace(newRow, "");
                if (text == "") continue;
                _documentData.Paragraflar.Add(text);
                _documentData.KelimeParagraflari.Add(paragraph);
                var style = paragraph.get_Style() as Word.Style;

                if (style?.NameLocal != "Başlık 2" || text.TrimStart().TrimEnd() != "KAYNAKLAR") continue;

                var pr = paragraph.Next();
                while (true)
                {
                    var resourceText = pr.Range.Text.TrimStart().TrimEnd().Replace(newLine, "").Replace(newRow, "");
                    if (resourceText != "")
                    {
                        if (resourceText?.IndexOf("[") > -1 && resourceText?.IndexOf("]") > -1)
                        {
                            if (resourceText?.TrimStart()?.TrimEnd()?.IndexOf("[", 0, 1) == 0)
                            {
                                resourceText = resourceText.TrimEnd().TrimStart();
                                resourceText = resourceText.Right(resourceText.Length - 4);
                            }
                            _documentData.Kaynaklar.Add(resourceText.Split(new[] { "[" }, StringSplitOptions.None)[0]);
                            _documentData.Kaynaklar.Add(resourceText.Split(new[] { "[" }, StringSplitOptions.None)[1].Split(new[] { "]" }, StringSplitOptions.None)[1]);
                        }
                        else
                        {
                            _documentData.Kaynaklar.Add(resourceText);
                        }
                    }
                    if (resourceText == "") break;
                    pr = pr.Next();
                }

            }
            #endregion

            #region Atıflar
            for (var i = 0; i < _documentData.Kaynaklar.Count; i++)
            {
                _documentData.Kaynaklar[i] = $"[{i + 1}] {_documentData.Kaynaklar[i].TrimStart().TrimEnd()}";
            }
            foreach (var resource in _documentData.Kaynaklar)
            {
                if (resource?.IndexOf("[") <= -1 || resource?.IndexOf("]") <= -1) continue;
                var resourceNum = resource?.Split(new[] { "[" }, StringSplitOptions.None)[1]
                    .Split(new[] { "]" }, StringSplitOptions.None)[0];
                resourceNum = $"[{resourceNum}]";
                var count = _documentData.Paragraflar.Where(x => x.Contains(resourceNum)).ToList().Count;
                if (count > 0)
                    _documentData.KaynakAtiflari.Add(resource);
                else
                    _documentData.HataliKaynakAtiflari.Add(resource);
            }
            #endregion

            await Task.Yield();
        }
        #endregion

        #region Tablolar ve Tablolara Atıflar
        private async Task Tablolar()
        {
            #region Tablolar
            foreach (Word.Paragraph paragraph in _wordDocument.Paragraphs)
            {
                var text = paragraph.Range.Text;
                text = text.TrimStart().TrimEnd().Replace(newLine, "").Replace(newRow, "");
                if (text == "") continue;
                _documentData.Paragraflar.Add(text);
                _documentData.KelimeParagraflari.Add(paragraph);
                var style = paragraph.get_Style() as Word.Style;
                if (text != "TABLOLAR LİSTESİ" || style?.NameLocal != "Normal") continue;
                var pw = paragraph;
                while (true)
                {
                    var tableName = pw.Next().Range.Text.TrimStart().TrimEnd().Replace(newLine, "").Replace(newRow, "");
                    pw = pw.Next();
                    if (tableName == "") break;
                    if (tableName?.IndexOf("...") == -1) continue;
                    tableName = tableName.Split(new[] { ". " }, StringSplitOptions.None)[0];
                    _documentData.Tablolar.Add(tableName);
                }
            }
            #endregion

            #region Atıflar
            foreach (var table in _documentData.Tablolar)
            {
                var useCount = _documentData.Paragraflar.Where(x => x.Contains(table) && !x.Contains("...")).ToList().Count;
                if (useCount > 0)
                    _documentData.TabloAtiflari.Add(table);
                else
                    _documentData.HataliTabloAtiflari.Add(table);
            }
            #endregion

            await Task.Yield();
        }
        #endregion

        #region Şekiller ve Şekillere Atıflar
        private async Task Sekiller()
        {
            #region Şekiller
            foreach (Word.Paragraph paragraph in _wordDocument.Paragraphs)
            {
                var text = paragraph.Range.Text;
                text = text.TrimStart().TrimEnd().Replace(newLine, "").Replace(newRow, "");
                if (text == "") continue;
                _documentData.Paragraflar.Add(text);
                _documentData.KelimeParagraflari.Add(paragraph);
                var style = paragraph.get_Style() as Word.Style;
                if (text != "ŞEKİLLER LİSTESİ" || style?.NameLocal != "Normal") continue;
                var pw = paragraph;
                while (true)
                {
                    var shapeName = pw.Next().Range.Text.TrimStart().TrimEnd().Replace(newLine, "").Replace(newRow, "");
                    pw = pw.Next();
                    if (shapeName == "") break;
                    if (shapeName?.IndexOf("...") == -1) continue;
                    shapeName = shapeName.Split(new[] { ". " }, StringSplitOptions.None)[0];
                    _documentData.Sekiller.Add(shapeName);
                }
            }
            #endregion

            #region Atıflar
            foreach (var shape in _documentData.Sekiller)
            {
                var useCount = _documentData.Paragraflar.Where(x => x.Contains(shape) && !x.Contains("...")).ToList().Count;
                if (useCount > 0)
                    _documentData.SekilAtiflari.Add(shape);
                else
                    _documentData.HataliSekilAtiflari.Add(shape);
            }
            #endregion

            await Task.Yield();
        }
        #endregion


    }
}