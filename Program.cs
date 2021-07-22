using System;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.IO;
using System.Data;
using System.Collections.Generic;


namespace DIO.Series 
{
    class Program
	
    {
		static SerieRepositorio repositorio = new SerieRepositorio();
		
		static void Main(string[] args)
        {
            // define nome do arquivo e pega o caminho da pasta onde ele esta
            // para verificar se ele existe senao ele cria um novo
            
            Diretorio(out string file, out string path);
			if (!File.Exists(path))
			{
				var wb = new XLWorkbook();
				var ws = wb.Worksheets.Add("Serie");
				
				ws.Cell("A1").Value = "Genero";
				ws.Cell("B1").Value = "Titulo";
				ws.Cell("C1").Value = "Descrição";
				ws.Cell("D1").Value = "Ano";
				ws.Cells("A1:E1").Style.Font.Bold = true;
				ws.Cells("A1:E1").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
				wb.SaveAs(path);
			}
			else
			{

			}
			string opcaoUsuario = ObterOpcaoUsuario();

            while (opcaoUsuario.ToUpper() != "X")
            {
                switch (opcaoUsuario)
                {
                    case "1":
                        ListarSeries();
                        break;
                    case "2":
                        InserirSerie();
                        break;
                    case "3":
                        AtualizarSerie();
                        break;
                    case "4":
                        ExcluirSerie();
                        break;
                    case "5":
                        VisualizarSerie();
                        break;
                    case "C":
                        Console.Clear();
                        break;

                    default:
                        throw new ArgumentOutOfRangeException();
                }

                opcaoUsuario = ObterOpcaoUsuario();
            }

            Console.WriteLine("Obrigado por utilizar nossos serviços.");
        }

        private static void Diretorio(out string file, out string path)
        {
			file = "serie.xlsx";
            path = Path.GetFullPath(file);
			
		}

        private static void ExcluirSerie()
		{
			Console.Write("Digite o id da série: ");
			int indiceSerie = int.Parse(Console.ReadLine());

			repositorio.Exclui(indiceSerie);
		}

        private static void VisualizarSerie()
		{
			Console.Write("Digite o id da série: ");
			int indiceSerie = int.Parse(Console.ReadLine());

			var serie = repositorio.RetornaPorId(indiceSerie);

			Console.WriteLine(serie);
		}

        private static void AtualizarSerie()
		{
			Console.Write("Digite o id da série: ");
			int indiceSerie = int.Parse(Console.ReadLine());

			// https://docs.microsoft.com/pt-br/dotnet/api/system.enum.getvalues?view=netcore-3.1
			// https://docs.microsoft.com/pt-br/dotnet/api/system.enum.getname?view=netcore-3.1
			foreach (int i in Enum.GetValues(typeof(Genero)))
			{
				Console.WriteLine("{0}-{1}", i, Enum.GetName(typeof(Genero), i));
			}
			Console.Write("Digite o gênero entre as opções acima: ");
			int entradaGenero = int.Parse(Console.ReadLine());

			Console.Write("Digite o Título da Série: ");
			string entradaTitulo = Console.ReadLine();

			Console.Write("Digite o Ano de Início da Série: ");
			int entradaAno = int.Parse(Console.ReadLine());

			Console.Write("Digite a Descrição da Série: ");
			string entradaDescricao = Console.ReadLine();

			Serie atualizaSerie = new Serie(id: indiceSerie,
										genero: (Genero)entradaGenero,
										titulo: entradaTitulo,
										ano: entradaAno,
										descricao: entradaDescricao);

			repositorio.Atualiza(indiceSerie, atualizaSerie);
		}
			private static void ListarSeries()
		{
			Console.WriteLine("Listar séries");

			var lista = repositorio.Lista();

			if (lista.Count == 0)
			{
				Console.WriteLine("Nenhuma série cadastrada.");
				return;
			}

			foreach (var serie in lista)
			{
                var excluido = serie.retornaExcluido();
                
				Console.WriteLine("#ID {0}: - Nome da série: {1} - Descrição: {2} {3}", serie.retornaId(), serie.retornaTitulo(), serie.retornaDescricao(),(excluido ? "*Excluído*" : ""));
			}
		}

        private static void InserirSerie()
		{
			Diretorio(out string file, out string path);
			
			Console.WriteLine("Inserir nova série");

			// https://docs.microsoft.com/pt-br/dotnet/api/system.enum.getvalues?view=netcore-3.1
			// https://docs.microsoft.com/pt-br/dotnet/api/system.enum.getname?view=netcore-3.1
			foreach (int i in Enum.GetValues(typeof(Genero)))
			{
				Console.WriteLine("{0}-{1}", i, Enum.GetName(typeof(Genero), i));
			}
			Console.Write("Digite o gênero entre as opções acima: ");
			int entradaGenero = int.Parse(Console.ReadLine());

			Console.Write("Digite o Título da Série: ");
			string entradaTitulo = Console.ReadLine();

			Console.Write("Digite o Ano de Início da Série: ");
			int entradaAno = int.Parse(Console.ReadLine());

			Console.Write("Digite a Descrição da Série: ");
			string entradaDescricao = Console.ReadLine();

			//Abrir arquivo excel e inserir os dados na planilha
			
			var wb = new XLWorkbook(path);
			var ws = wb.Worksheet(1);
			
			//Verfica a coluna para ver qual ultima linha que tem dados para inserir novo abaixo 
			var UltimaLinha2 = ws.Column(1).LastCellUsed().Address.RowNumber;
			var UltimaLinha3 = ws.Column(2).LastCellUsed().Address.RowNumber;
			var UltimaLinha4 = ws.Column(3).LastCellUsed().Address.RowNumber;
			var UltimaLinha5 = ws.Column(4).LastCellUsed().Address.RowNumber;

			//Insere os dados na planilha
			ws.Cell("a" + UltimaLinha2).CellBelow(1).Value = (Genero)entradaGenero;
			ws.Cell("b" + UltimaLinha3).CellBelow(1).Value = entradaTitulo;
			ws.Cell("c" + UltimaLinha4).CellBelow(1).Value = entradaDescricao;
			ws.Cell("d" + UltimaLinha5).CellBelow(1).Value = entradaAno;

			ws.Columns("a:d").AdjustToContents().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
				
			wb.SaveAs(path);

		}

		private static string ObterOpcaoUsuario()
		{
			Console.WriteLine();
			Console.WriteLine("DIO Séries a seu dispor!!!");
			Console.WriteLine("Informe a opção desejada:");

			Console.WriteLine("1- Listar séries");
			Console.WriteLine("2- Inserir nova série");
			Console.WriteLine("3- Atualizar série");
			Console.WriteLine("4- Excluir série");
			Console.WriteLine("5- Visualizar série");
			Console.WriteLine("C- Limpar Tela");
			Console.WriteLine("X- Sair");
			Console.WriteLine();

			string opcaoUsuario = Console.ReadLine().ToUpper();
			Console.WriteLine();
			return opcaoUsuario;
		}
    }

}
