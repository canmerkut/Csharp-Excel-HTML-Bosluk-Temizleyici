using System;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using System.IO;

namespace ExcelRemoveHTMLSpaces
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // EPPlus lisans ayarı
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string dosyaYolu = "urunler0.xls";
            string yeniDosyaYolu = "urunler_temiz.xlsx";

            try
            {
                using (var package = new ExcelPackage(new FileInfo(dosyaYolu)))
                {
                    var worksheet = package.Workbook.Worksheets[0]; // İlk sayfayı al
                    var son = worksheet.Dimension.End;

                    // 9. sütundaki tüm satırları işle
                    for (int row = 1; row <= son.Row; row++)
                    {
                        var hucre = worksheet.Cells[row, 9].Text;
                        if (!string.IsNullOrEmpty(hucre))
                        {
                            // Birden fazla boşluğu tek boşluğa çevir
                            string temizMetin = Regex.Replace(hucre, @"\s+", " ").Trim();
                            worksheet.Cells[row, 9].Value = temizMetin;
                        }
                    }

                    // Yeni dosyaya kaydet
                    package.SaveAs(new FileInfo(yeniDosyaYolu));
                }

                Console.WriteLine("İşlem başarıyla tamamlandı!");
                Console.WriteLine($"Yeni dosya kaydedildi: {yeniDosyaYolu}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Hata oluştu: {ex.Message}");
            }

            Console.WriteLine("Çıkmak için bir tuşa basın...");
            Console.ReadKey();
        }
    }
}
