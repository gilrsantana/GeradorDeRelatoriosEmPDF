using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace GeradorDeRelatoriosEmPDF
{
    class Program
    {
        static List<Pessoa> pessoas = new List<Pessoa>();
        static BaseFont fonteBase = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);
        static void Main(string[] args)
        {
            DesserializarPessoas();
            GerarRelatorioEmPDF(1000);
        }

        static void DesserializarPessoas()
        {
            if(File.Exists("pessoas.json"))
            {
                using(var sr = new StreamReader("pessoas.json"))
                {
                    var dados = sr.ReadToEnd();
                    pessoas = JsonSerializer.Deserialize(dados, typeof(List<Pessoa>)) as List<Pessoa>;
                }
            }
        }

        static void GerarRelatorioEmPDF(int qtdePessoas)
        {
            var pessoasSelecionadas = pessoas.Take(qtdePessoas).ToList();
            if(pessoasSelecionadas.Count > 0)
            {
                // Cálculo da quantidade total de linhas
                int totalDePaginas = 1;
                int totalDeLinhas = pessoasSelecionadas.Count;
                if (totalDeLinhas > 24)
                    totalDePaginas += (int)Math.Ceiling((totalDeLinhas - 24) / 29F);

                // Configuração do Documento PDF
                int dpi = 72; // densidade padrão é de 72 dpi
                var polegada = 25.2F;
                var pxPorMm = dpi / polegada;
                var pdf = new Document(PageSize.A4, 15 * pxPorMm, 15 * pxPorMm, 15 * pxPorMm, 20 * pxPorMm);
                var nomeArquivo = $"pessoas.{DateTime.Now.ToString("yyyy.MM.dd.HH.mm.ss")}.pdf";
                var arquivo = new FileStream(nomeArquivo, FileMode.Create);
                var writer = PdfWriter.GetInstance(pdf, arquivo);
                writer.PageEvent = new EventosDePagina(totalDePaginas);
                pdf.Open();

                //Adição do título
                var fonteParagrafo = new iTextSharp.text.Font(fonteBase, 32, iTextSharp.text.Font.NORMAL, BaseColor.Black);
                var titulo = new Paragraph("Relatorio de Pessoas\n\n", fonteParagrafo);
                titulo.Alignment = Element.ALIGN_LEFT;
                titulo.SpacingAfter = 4;
                pdf.Add(titulo);

                // Adição da Imagem
                var caminhoImagem = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "img\\youtube.png");
                if (File.Exists(caminhoImagem))
                {
                    iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(caminhoImagem);
                    float razaoAlturaLargura = logo.Width / logo.Height; // Calcula a relação de proporção da imagem
                    float alturaLogo = 32;
                    float larguraLogo = alturaLogo * razaoAlturaLargura;
                    logo.ScaleToFit(larguraLogo, alturaLogo);
                    var margemEsquerda = pdf.PageSize.Width - pdf.RightMargin - larguraLogo;
                    var margemTopo = pdf.PageSize.Height - pdf.TopMargin - 54;
                    logo.SetAbsolutePosition(margemEsquerda, margemTopo);
                    writer.DirectContent.AddImage(logo, false);
                }

                // Adição de Link
                var fonteLink = new iTextSharp.text.Font(fonteBase, 9.9F, Font.NORMAL, BaseColor.Blue);
                var link = new Chunk("Canal do Professor Ricardo Maroquio", fonteLink);
                link.SetAnchor("https://www.youtube.com/maroquio");
                var larguraTextoLink = fonteBase.GetWidthPoint(link.Content, fonteLink.Size);
                var caixaTexto = new ColumnText(writer.DirectContent);
                caixaTexto.AddElement(link);
                caixaTexto.SetSimpleColumn(
                    pdf.PageSize.Width - pdf.RightMargin - larguraTextoLink,
                    pdf.PageSize.Height - pdf.TopMargin - (30 * pxPorMm),
                    pdf.PageSize.Width - pdf.RightMargin,
                    pdf.PageSize.Height - pdf.TopMargin - (18 * pxPorMm)
                    );
                caixaTexto.Go();

                // Adicionar Tabela
                var tabela = new PdfPTable(5);
                float[] larguraColunas = { 0.6f, 2f, 1.5f, 1f, 1f };
                tabela.DefaultCell.BorderWidth = 0;
                tabela.WidthPercentage = 100;

                // Adição das células de títulos das colunas
                CriarCelulaTexto(tabela, "Código", PdfCell.ALIGN_CENTER, true);
                CriarCelulaTexto(tabela, "Nome", PdfCell.ALIGN_LEFT, true);
                CriarCelulaTexto(tabela, "Profissão", PdfCell.ALIGN_CENTER, true);
                CriarCelulaTexto(tabela, "Salário", PdfCell.ALIGN_CENTER, true);
                CriarCelulaTexto(tabela, "Empregada", PdfCell.ALIGN_CENTER, true);

                foreach (var pessoa in pessoasSelecionadas)
                {
                    CriarCelulaTexto(tabela, pessoa.IdPessoa.ToString("D6"), PdfCell.ALIGN_CENTER);
                    CriarCelulaTexto(tabela, pessoa.Nome + " " + pessoa.Sobrenome, PdfCell.ALIGN_LEFT);
                    CriarCelulaTexto(tabela, pessoa.Profissao.Nome, PdfCell.ALIGN_CENTER, true);
                    CriarCelulaTexto(tabela, pessoa.Salario.ToString("C2"), PdfCell.ALIGN_RIGHT);
                    //CriarCelulaTexto(tabela, pessoa.Empregado ? "Sim" : "Não", PdfCell.ALIGN_CENTER);
                    var caminhoImagemCelula = pessoa.Empregado ? "img\\emoji_feliz.png" : "img\\emoji_triste.png";
                    caminhoImagemCelula = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, caminhoImagemCelula);
                    CriarCelulaImagem(tabela, caminhoImagemCelula, 20, 20);
                }

                pdf.Add(tabela);

                pdf.Close();
                arquivo.Close();

                // Abre o PDF pelo visualizador padrão do sistema
                var caminhoPDF = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, nomeArquivo);
                if (File.Exists(caminhoPDF))
                {
                    Process.Start(new ProcessStartInfo()
                    {
                        Arguments = $"/c start {caminhoPDF}", // Abre pelo visualizador padrão do sistema
                        //Arguments = $"/c start firefox {caminhoPDF}", Abre pelo firefox
                        //Arguments = $"/c start chrome {caminhoPDF}", Abre pelo Chrome
                        //Arguments = $"/c start acrobat {caminhoPDF}", Abre pelo Acrobat Reader
                        FileName = "cmd.exe",
                        CreateNoWindow = true
                    });
                }
            }
        }

        private static void CriarCelulaTexto(PdfPTable tabela, 
            string texto, 
            int alinhamentoHorz = PdfPCell.ALIGN_LEFT, 
            bool negrito = false, 
            bool italico = false,
            int tamanhoFonte = 12,
            int alturaCelula = 25)
        {
            int estilo = iTextSharp.text.Font.NORMAL;
            if (negrito && italico)
            {
                estilo = iTextSharp.text.Font.BOLDITALIC;
            }
            else if (negrito)
            {
                estilo = iTextSharp.text.Font.BOLD;
            }
            else if (italico)
            {
                estilo = iTextSharp.text.Font.ITALIC;
            }
            var fonteCelula = new iTextSharp.text.Font(fonteBase, tamanhoFonte, estilo, BaseColor.Black);
            var bgColor = iTextSharp.text.BaseColor.White;
            if (tabela.Rows.Count % 2 == 1)
                bgColor = new BaseColor(0.95F, 0.95F, 0.95F);
            var celula = new PdfPCell(new Phrase(texto, fonteCelula));
            celula.HorizontalAlignment = alinhamentoHorz;
            celula.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
            celula.Border = 0;
            celula.BorderWidthBottom = 1;
            celula.FixedHeight = alturaCelula;
            celula.PaddingBottom = 5;
            celula.BackgroundColor = bgColor;
            tabela.AddCell(celula);
        }

        static void CriarCelulaImagem(PdfPTable tabela, string caminhoImagem, int larguraImagem, int alturaImagem, int alturaCelula = 25)
        {
            var bgColor = iTextSharp.text.BaseColor.White;
            if (tabela.Rows.Count % 2 == 1)
                bgColor = new BaseColor(0.95F, 0.95F, 0.95F);
            if(File.Exists(caminhoImagem))
            {
                iTextSharp.text.Image imagem = iTextSharp.text.Image.GetInstance(caminhoImagem);
                imagem.ScaleToFit(larguraImagem, alturaImagem);
                var celula = new PdfPCell(imagem);
                celula.HorizontalAlignment = PdfCell.ALIGN_CENTER;
                celula.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                celula.Border = 0;
                celula.BorderWidthBottom = 1;
                celula.FixedHeight = alturaCelula;
                celula.BackgroundColor = bgColor;
                tabela.AddCell(celula);
            }
        }
    }
}
