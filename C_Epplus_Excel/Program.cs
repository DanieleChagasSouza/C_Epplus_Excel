using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;

class Program
{
    static void Main(string[] args)
    {
        // Define o caminho onde o arquivo Excel será salvo
        string filePath = "C:\\Users\\DanieleChagasSouza\\Desktop\\Excel Epplus\\arquivo.xlsx";

        Console.WriteLine("Pressione algo para iniciar!");
        Console.ReadKey();
        CriarPlanilhaExcel(filePath);

        Console.WriteLine("Pressione algo para acessar e exibir a planilha\n");
        Console.ReadKey();
        AbrePlanilhaExcel(filePath);

        Console.ReadKey();
    }
    //O metodo que criar a planilha excel 
    private static void CriarPlanilhaExcel(string caminhoPlanilha)
    {
        //  Define uma lista de objetos anônimos que serão os dados da planilha.
        var planilhas = new[]
        {
            new {Id = "SP01", Nome = "João", Idade = 29},
            new {Id = "RJ02", Nome = "Daiane", Idade = 25},
            new {Id = "SC03", Nome = "Maria", Idade = 31},
            new {Id = "MG04", Nome = "Daniele", Idade = 32},
            new {Id = "SG05", Nome = "Ananias", Idade = 30},
            new {Id = "BA06", Nome = "José", Idade = 33}
        };

        // Define o contexto de licença para uso não comercial do EPPlus (NonCommercial),
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        //Cria uma nova instância do pacote Excel.
        using (ExcelPackage excel = new ExcelPackage())
        {
            // Adiciona uma nova planilha ao pacote Excel.
            var worksheet = excel.Workbook.Worksheets.Add("Planilha1");

            // Define a cor da aba da planilha.
            worksheet.TabColor = System.Drawing.Color.Black;

            //Define a altura padrão das linhas
            worksheet.DefaultRowHeight = 12;

            // Define a altura da primeira linha.
            worksheet.Row(1).Height = 30;

            //Alinha o texto da primeira linha ao centro.
            worksheet.Row(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            //Define o texto da primeira linha como negrito.
            worksheet.Row(1).Style.Font.Bold = true;

            // Define o valor da célula na linha 1, coluna 1, 2,3.
            worksheet.Cells[1, 1].Value = "Cod";
            worksheet.Cells[1, 2].Value = "Nome";
            worksheet.Cells[1, 3].Value = "Idade/int";

            //Define o texto das células no intervalo "A1" como itálico.
            worksheet.Cells["A1:C1"].Style.Font.Italic = true;

            // Adiciona cor de fundo ao cabeçalho
            worksheet.Cells["A1:C1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells["A1:C1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);


            // Itera sobre os dados e preenche as células da planilha.
            int indice = 2;
            foreach (var planilha in planilhas)
            {
                worksheet.Cells[indice, 1].Value = planilha.Id;
                worksheet.Cells[indice, 2].Value = planilha.Nome;
                worksheet.Cells[indice, 3].Value = planilha.Idade;
                indice++;
            }

            // Ajusta automaticamente a largura da coluna 1, 2, 3.
            worksheet.Column(1).AutoFit();
            worksheet.Column(2).AutoFit();
            worksheet.Column(3).AutoFit();

            // Se o arquivo existir, exclui
            if (File.Exists(caminhoPlanilha)) File.Delete(caminhoPlanilha);

            // Cria o arquivo excel no disco físico
            File.WriteAllBytes(caminhoPlanilha, excel.GetAsByteArray());
        }

        Console.WriteLine($"Planilha criada com sucesso: {caminhoPlanilha} \n");
    }
    //O metodo que Abre a planilha excel.
    private static void AbrePlanilhaExcel(string caminhoPlanilha)
    {
        // Abre o arquivo Excel existente.
        using (var arquivoExcel = new ExcelPackage(new FileInfo(caminhoPlanilha)))
        {
            // Obtém a primeira planilha do pacote Excel.
            ExcelWorksheet planilha = arquivoExcel.Workbook.Worksheets.FirstOrDefault();

            //Verifica se a planilha foi encontrada.
            if (planilha == null)
            {
                Console.WriteLine("Nenhuma planilha encontrada!");
                return;
            }

            // Obter o número de linhas e colunas
            int rows = planilha.Dimension.Rows;
            int cols = planilha.Dimension.Columns;

            // Percorre as linhas e colunas da planilha
            for (int i = 1; i <= rows; i++)
            {
                for (int j = 1; j <= cols; j++)
                {
                    string conteudo = planilha.Cells[i, j].Value?.ToString() ?? string.Empty;
                    Console.WriteLine(conteudo); //Obtém o valor da célula e exibe no console.
                }
            }
        }
    }
}
