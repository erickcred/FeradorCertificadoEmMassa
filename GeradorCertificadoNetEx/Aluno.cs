using System;

namespace GeradorCertificadoNetEx
{
    [Serializable]
    public class Aluno
    {
        public string Nome { get; set; }
        public string Cpf { get; set; }
        public string Crc { get; set; }
        public string Email { get; set; }


        public string CursoNome { get; set; }
        public string Codigo { get; set; }
        public string Categorias { get; set; }
        public double Pontos { get; set; }
        public double CargaHoraria { get; set; }
        public double Nota { get; set; }
        public double Frequencia { get; set; }
        public string DataConclusao { get; set; }

    }
}