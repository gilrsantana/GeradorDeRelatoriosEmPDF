using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeradorDeRelatoriosEmPDF
{
    class EventosDePagina : PdfPageEventHelper
    {
        private BaseFont fonteBaseRodape { get; set; }
        private iTextSharp.text.Font fonteRodape { get; set; }
        public int TotalDePaginas { get; set; }

        private PdfContentByte wdc;

        public EventosDePagina(int totalDePaginas)
        {
            fonteBaseRodape = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);
            fonteRodape = new iTextSharp.text.Font(fonteBaseRodape, 8f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.Black);
            TotalDePaginas = totalDePaginas;
        }

        public override void OnOpenDocument(PdfWriter writer, Document document)
        {
            base.OnOpenDocument(writer, document);

            this.wdc = writer.DirectContent;
        }
        public override void OnEndPage(PdfWriter writer, Document document)
        {
            base.OnEndPage(writer, document);

            AdicionarMomentoGeracaoRelatorio(writer, document);
            AdicionarNumeroDePaginas(writer, document);
        }

        private void AdicionarMomentoGeracaoRelatorio(PdfWriter writer, Document document)
        {
            var textMomentoGeracao = $"Gerado em {DateTime.Now.ToShortDateString()} às {DateTime.Now.ToShortTimeString()}";
            wdc.BeginText();
            wdc.SetFontAndSize(fonteRodape.BaseFont, fonteRodape.Size);
            wdc.SetTextMatrix(document.LeftMargin, document.BottomMargin * 0.75f);
            wdc.ShowText(textMomentoGeracao);
            wdc.EndText();
        }

        private void AdicionarNumeroDePaginas(PdfWriter writer, Document document)
        {
            int paginaAtual = writer.PageNumber;
            var textoPaginacao = $"Página {paginaAtual} de {TotalDePaginas}";
            float larguraTextoPaginacao = fonteBaseRodape.GetWidthPoint(textoPaginacao, fonteRodape.Size);
            var tamanhoPagina = document.PageSize;
            wdc.BeginText();
            wdc.SetFontAndSize(fonteRodape.BaseFont, fonteRodape.Size);
            wdc.SetTextMatrix(tamanhoPagina.Width - document.RightMargin - larguraTextoPaginacao, document.BottomMargin * 0.75f);
            wdc.ShowText(textoPaginacao);
            wdc.EndText();
        }
    }
}
