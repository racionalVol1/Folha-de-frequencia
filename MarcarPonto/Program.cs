public class Program
{    
    private const string CaminhoDaPastaDeRegistros = @"C:\Users\p051605\OneDrive - rede.sp\Área de Trabalho\Folha de frequencia\registros";

    public static void Main(string[] args)
    {
        if (args.Length == 0 || (args[0].ToLower() != "entrada" && args[0].ToLower() != "saida"))
        {
            Console.WriteLine("Erro: Use o comando da seguinte forma:");
            Console.WriteLine("dotnet run entrada");
            Console.WriteLine("ou");
            Console.WriteLine("dotnet run saida");
            return;
        }

        string evento = args[0].ToLower();
        RegistrarPonto(evento);
    }

    public static void RegistrarPonto(string evento)
    {
        DateTime agora = DateTime.Now;
        string dataFormatada = agora.ToString("dd/MM/yyyy");
        string horaFormatada = agora.ToString("HH:mm");
        string nomeArquivo = $"registros_{agora:MM-yyyy}.txt";

        try
        {            
            Directory.CreateDirectory(CaminhoDaPastaDeRegistros);

            
            string caminhoCompleto = Path.Combine(CaminhoDaPastaDeRegistros, nomeArquivo);

            string linhaRegistro = $"{dataFormatada},{evento},{horaFormatada}{Environment.NewLine}";
            File.AppendAllText(caminhoCompleto, linhaRegistro);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Ponto de '{evento}' registrado com sucesso!");
            Console.WriteLine($"Salvo em: {caminhoCompleto}");
            Console.ResetColor();
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Ocorreu um erro ao salvar o registro: {ex.Message}");
            Console.ResetColor();
        }
    }
}