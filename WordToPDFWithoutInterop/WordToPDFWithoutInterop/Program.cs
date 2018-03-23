using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordToPDF;
namespace WordToPDFWithoutInterop
{
    class Program
    {
        static void Main(string[] args)
        {
            Word2Pdf ObjetWord = new Word2Pdf();
            string dossier = "Z:\\projet tutore\\";
            string nomDuFichierAConvertir = "cv.doc"; //ou docx
            object CheminDuFichier = dossier + "\\" + nomDuFichierAConvertir;
            string ExtensionDuFichier = Path.GetExtension(nomDuFichierAConvertir);
            string ExtensionCible = nomDuFichierAConvertir.Replace(ExtensionDuFichier, ".pdf");
            if (ExtensionDuFichier == ".doc" || ExtensionDuFichier == ".docx")
            {
                object DossierCible = dossier + "\\" + ExtensionCible;
                ObjetWord.InputLocation = CheminDuFichier;
                ObjetWord.OutputLocation = DossierCible;
                ObjetWord.Word2PdfCOnversion();
            }
        }
    }
}
