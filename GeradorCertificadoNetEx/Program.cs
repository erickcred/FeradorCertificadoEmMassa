using System;
using System.Text.Json;
using System.Linq;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using ClosedXML.Excel;
using System.Diagnostics;

namespace GeradorCertificadoNetEx
{
    public class Program
    {
        private static List<Aluno> Alunos = new List<Aluno>();

        public static void Main(string[] args)
        {
            // DesserializarAlunos();
            // foreach (var aluno in Alunos)
            // {
            //     Console.WriteLine(aluno.Nome);
            // }

            LerExcell();
            GerarCertificado(Alunos);


        }

        public static void DesserializarAlunos()
        {
            if (File.Exists("aluno.json"))
            {
                using (var seria = new StreamReader("aluno.json"))
                {
                    var dados = seria.ReadToEnd();
                    Alunos = JsonSerializer.Deserialize(dados, typeof(List<Aluno>)) as List<Aluno>;
                }
            }
        }

        public static void LerExcell()
        {
            //var xlsx = new XLWorkbook(@"C:\Users\erick.oliveira\Documents\IOBEducacao\Projetos\GeradorCertificadoNetEx\GeradorCertificadoNetEx\aluno.xlsx");
            //var xlsx = new XLWorkbook(@"C:\aluno.xlsx");
            var xlsx = new XLWorkbook(@"C:\Users\erick.oliveira\Desktop\aluno.xlsx");
            var planilha = xlsx.Worksheets.First(w => w.Name == "Planilha1");
            var totalLinhas = planilha.Rows().Count();

            for (int i = 2; i < totalLinhas + 1; i++)
            {
                string email;
                string nome;
                string cpf;
                string crc;
                string cursoNome;
                string codigo;
                string categorias;
                string dataConclusao;
                double pontos, cargaHoraria, nota, frequencia;

                if (planilha.Cell($"B{i}").Value.ToString() == null) { nome = ""; }
                else { nome = planilha.Cell($"B{i}").Value.ToString(); }

                if (planilha.Cell($"C{i}").Value.ToString() == null) { cpf = " -- "; }
                else { cpf = planilha.Cell($"C{i}").Value.ToString(); }

                if (planilha.Cell($"D{i}").Value.ToString() == null) { crc = " -- "; }
                else { crc = planilha.Cell($"D{i}").Value.ToString(); }

                if (planilha.Cell($"E{i}").Value.ToString() == null) { cursoNome = " -- "; }
                else { cursoNome = planilha.Cell($"E{i}").Value.ToString(); }

                if (planilha.Cell($"F{i}").Value.ToString() == null) { codigo = " -- "; }
                else { codigo = planilha.Cell($"F{i}").Value.ToString(); }

                if (planilha.Cell($"G{i}").Value.ToString() == null) { categorias = " -- "; }
                else { categorias = planilha.Cell($"G{i}").Value.ToString(); }

                if (planilha.Cell($"H{i}").Value.ToString() == null) { pontos = 0; }
                else { pontos = Convert.ToDouble(planilha.Cell($"H{i}").Value.ToString()); }

                if (planilha.Cell($"I{i}").Value.ToString() == null) { cargaHoraria = 0; }
                else { cargaHoraria = Convert.ToDouble(planilha.Cell($"I{i}").Value.ToString()); }

                if (planilha.Cell($"J{i}").Value.ToString() == null) { nota = 0; }
                else { nota = Convert.ToDouble(planilha.Cell($"J{i}").Value.ToString()); }

                if (planilha.Cell($"K{i}").Value.ToString() == null) { frequencia = 0; }
                else { frequencia = Convert.ToDouble(planilha.Cell($"K{i}").Value.ToString()); }

                if (planilha.Cell($"L{i}").Value.ToString() == null) { dataConclusao = ""; }
                else { dataConclusao = planilha.Cell($"L{i}").Value.ToString().Split(" ")[0]; }

                try
                {
                    Alunos.Add(new Aluno
                    {
                        Email = planilha.Cell($"A{i}").Value.ToString(),
                        Nome = nome,
                        Cpf = cpf,
                        Crc = crc,
                        CursoNome = cursoNome,
                        Codigo = codigo,
                        Categorias = categorias,
                        Pontos = pontos,
                        CargaHoraria = cargaHoraria,
                        Nota = nota,
                        Frequencia = frequencia,
                        DataConclusao = dataConclusao
                    });
                } catch (FormatException error)
                {
                    Console.WriteLine(error.Message);
                }
            }
        }

        public static void GerarCertificado(List<Aluno> alunos)
        {
            foreach (var a in Alunos)
            {
                //var aluno = Alunos.Take(quantidade).ToList();
                if (alunos.Count > 0)
                {
                    // Configuração do PDF
                    var pxPorMm = 72 / 25.2f;
                    var pdf = new Document(PageSize.A4, 25 * pxPorMm, 25 * pxPorMm, 0 * pxPorMm, 0 * pxPorMm);
                    pdf.SetPageSize(PageSize.A4.Rotate());

                    var nomeAluno = "";
                    if (a.Nome.Length <= 15)
                    {
                        nomeAluno = a.Nome;
                    } else
                    {
                        nomeAluno = a.Nome.Substring(0, 15);
                    }

                    var nomeArquivo = @$"C:\Certificados\{a.CursoNome.Substring(0, 10)}_{a.Cpf}_{nomeAluno}_{DateTime.Now.ToString("dd.MM.yyyy_HH.mm")}.pdf";
                    var arquivo = new FileStream(nomeArquivo, FileMode.Create);

                    var pdfWriter = PdfWriter.GetInstance(pdf, arquivo); // associando o arquivo ao pdf
                    pdf.Open();

                    var fontBase = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);
                    var fontParagrafo = new Font(fontBase, 13, Font.NORMAL, BaseColor.Black);
                    var fontNegrito = new Font(fontBase, 13, Font.BOLD, BaseColor.Blue);
                    var fontDestaque = new Font(fontBase, 13, Font.BOLD, new BaseColor(92, 0, 117));

                    Paragraph descricao;
                    if ((a.Codigo.Equals("") && a.Categorias.Equals("")) || a.Codigo.Equals("") || a.Categorias.Equals(""))
                    {
                        descricao = new Paragraph(@$"Certificamos que o(a) Sr(a) {a.Nome}, portador(a) do CRC de nº {a.Crc}, concluiu em {a.DataConclusao} o curso {a.CursoNome}, com a carga horária de {a.CargaHoraria} horas, e atingiu o critério de avaliação exigido para aprovação.", fontParagrafo);
                    } else
                    {
                        descricao = new Paragraph(@$"Certificamos que o(a) Sr(a) {a.Nome}, portador(a) do CRC de nº {a.Crc}, concluiu em {a.DataConclusao} o curso {a.CursoNome}, código {a.Codigo}, com a carga horária de {a.CargaHoraria} horas, e atingiu o critério de avaliação exigido para aprovação.", fontParagrafo);
                    }
                    descricao.Alignment = Element.ALIGN_LEFT;
                    pdf.Add(new Paragraph("\n\n\n\n\n\n\n\n\n\n\n\n\n"));
                    pdf.Add(descricao);

                    Paragraph categorias;
                    if ((!a.Codigo.Equals("") && !a.Categorias.Equals("")) || !a.Codigo.Equals("") || !a.Categorias.Equals(""))
                    {
                        categorias = new Paragraph($"\nPontuação CFC: {a.Pontos} pontos nas categorias {a.Categorias}.", fontDestaque);
                        pdf.Add(categorias);
                    }

                    var data = new Paragraph($"\nSão Paulo, {DateTime.Now.Day} de {DateTime.Now.ToString("Y")}", fontParagrafo);
                    data.Alignment = Element.ALIGN_CENTER;
                    pdf.Add(data);




                    // Imagem
                    var caminhoImagemModelo = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @$"img\ModeloCertificado.png");
                    if (File.Exists(caminhoImagemModelo))
                    {
                        Image modelo = Image.GetInstance(caminhoImagemModelo);
                        float alturaLargura = modelo.Width / modelo.Height;
                        float altura = 845;
                        float largura = altura * alturaLargura;
                        modelo.ScaleToFit(altura, largura);

                        var margemEsquerda = 0;
                        var margemTopo = 0;
                        modelo.SetAbsolutePosition(margemEsquerda, margemTopo);
                        //pdfWriter.DirectContent.AddImage(modelo, false);
                        pdf.Add(modelo);
                    }



                    pdf.Close();
                    arquivo.Close();

                    // Abrir Pdf
                    var caminhoPdf = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, nomeArquivo + ".pdf");
                    if (File.Exists(caminhoPdf))
                    {
                        Process.Start(new ProcessStartInfo()
                        {
                            Arguments = $"/c start {caminhoPdf}",
                            FileName = "cmd.exe",
                            CreateNoWindow = true
                        });
                    }
                    Console.Clear();
                    Console.WriteLine("caminho: " + caminhoImagemModelo);
                    Console.ForegroundColor = ConsoleColor.Blue;
                    Console.WriteLine($"Certificados Gerados com sucesso! em {caminhoPdf}");
                    Console.ForegroundColor = ConsoleColor.White;

                } else
                {
                    Console.WriteLine(@"Não foi encontrado na em C:\  o arquivo (aluno.xlsx) ou ele não está no formato correto! Porfavor verifique o arquivo");
                    Console.ReadKey();
                }

            }
        }
    }
}