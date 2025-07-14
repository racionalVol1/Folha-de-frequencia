// --- VERSÃO FINAL E ESPECÍFICA ---
using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using Xceed.Words.NET;

public class Program
{
    private const string CaminhoDaPastaDeRegistros = @"C:\Users\p051605\OneDrive - rede.sp\Área de Trabalho\Folha de frequencia\registros";

    public static void Main(string[] args)
    {
        string arquivoModelo = "Folha de frequência Modelo.docx";
        
        // Pega o mês e ano atuais para usar como padrão.
        DateTime dataAtual = DateTime.Now;
        string mesAtual = dataAtual.ToString("MM");
        string anoAtual = dataAtual.ToString("yyyy");
        
        string nomeArquivoRegistros = $"registros_{mesAtual}-{anoAtual}.txt";
        string caminhoCompletoRegistros = Path.Combine(CaminhoDaPastaDeRegistros, nomeArquivoRegistros);
        string mesAnoSaida = $"{mesAtual}-{anoAtual}";

        PreencherFolhaPonto(arquivoModelo, caminhoCompletoRegistros, mesAnoSaida);
    }

    private static string GetDiaSemana(string dataStr)
    {
        try
        {
            DateTime dataObj = DateTime.ParseExact(dataStr, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            var cultura = new CultureInfo("pt-BR");
            string diaDaSemana = cultura.DateTimeFormat.GetDayName(dataObj.DayOfWeek);
            return $"{char.ToUpper(diaDaSemana[0])}{diaDaSemana.Substring(1)}" + ( (int)dataObj.DayOfWeek > 0 && (int)dataObj.DayOfWeek < 6 ? "-feira" : "" );
        }
        catch { return ""; }
    }

    public static void PreencherFolhaPonto(string modeloPath, string registrosPath, string mesAnoSaida)
    {
        if (!File.Exists(registrosPath))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"ERRO: Arquivo de registros não encontrado em '{registrosPath}'");
            Console.ResetColor();
            return;
        }

        var pontos = new Dictionary<int, Dictionary<string, string>>();
        var linhasRegistros = File.ReadAllLines(registrosPath);

        foreach (var linha in linhasRegistros)
        {
            if (string.IsNullOrWhiteSpace(linha)) continue; // Ignora linhas em branco
            var partes = linha.Split(',');
            if (partes.Length == 3)
            {
                int dia = int.Parse(partes[0].Split('/')[0]);
                string evento = partes[1];
                string hora = partes[2];
                if (!pontos.ContainsKey(dia)) { pontos[dia] = new Dictionary<string, string>(); }
                pontos[dia][evento] = hora;
            }
        }

        try
        {
            using (var doc = DocX.Load(modeloPath))
            {
                if (doc.Tables.Count == 0) { Console.WriteLine("ERRO: Nenhuma tabela encontrada."); return; }

                // ===================================================================
                // <<< ALTERAÇÃO 1: MÊS/ANO >>>
                // Vamos procurar a célula que contém "MÊS/ANO" e alterar a célula ao lado dela.
                // ===================================================================
                bool mesAnoAlterado = false;
                foreach (var table in doc.Tables)
                {
                    foreach (var row in table.Rows)
                    {
                        for (int i = 0; i < row.Cells.Count; i++)
                        {
                            // Se encontrarmos o texto "MÊS/ANO" em uma célula...
                            if (row.Cells[i].Paragraphs[0].Text.Contains("MÊS/ANO"))
                            {
                                // ...alteramos a célula seguinte na mesma linha.
                                if (i + 1 < row.Cells.Count)
                                {
                                    var celulaMesAno = row.Cells[i + 1];
                                    string mesFormatado = mesAnoSaida.Replace("-", "/").ToUpper(); // Formato: "07/2025"
                                    celulaMesAno.Paragraphs[0].RemoveText(0);
                                    celulaMesAno.Paragraphs[0].Append(mesFormatado).Bold(); // Deixa em negrito como no modelo
                                    mesAnoAlterado = true;
                                    break;
                                }
                            }
                        }
                        if (mesAnoAlterado) break;
                    }
                    if (mesAnoAlterado) break;
                }


                // ===================================================================
                // <<< ALTERAÇÃO 2 e 3: PERÍODO e OBSERVAÇÃO >>>
                // Focando na tabela principal.
                // ===================================================================
                var tabelaPrincipal = doc.Tables[0]; // Assumindo que a tabela principal é a primeira
                for (int i = 2; i < tabelaPrincipal.Rows.Count; i++) // Começa em i=2 para pular cabeçalhos
                {
                    var linha = tabelaPrincipal.Rows[i];
                    string diaTabelaStr = linha.Cells[0].Paragraphs[0].Text.Trim();

                    if (int.TryParse(diaTabelaStr, out int diaTabela) && pontos.ContainsKey(diaTabela))
                    {
                        var dadosDia = pontos[diaTabela];
                        string dataCompleta = $"{diaTabela:D2}/{mesAnoSaida.Split('-')[0]}/{mesAnoSaida.Split('-')[1]}";

                        // Coluna de ENTRADA (índice 1)
                        if (dadosDia.ContainsKey("entrada"))
                        {
                            linha.Cells[1].Paragraphs[0].RemoveText(0);
                            linha.Cells[1].Paragraphs[0].Append(dadosDia["entrada"]);
                        }

                        // Coluna de SAÍDA (índice 2)
                        if (dadosDia.ContainsKey("saida"))
                        {
                            linha.Cells[2].Paragraphs[0].RemoveText(0);
                            linha.Cells[2].Paragraphs[0].Append(dadosDia["saida"]);
                        }

                        // Coluna de OBSERVAÇÃO (índice 3)
                        linha.Cells[3].Paragraphs[0].RemoveText(0);
                        linha.Cells[3].Paragraphs[0].Append(GetDiaSemana(dataCompleta));
                    }
                }
                
                string nomeArquivoSaida = $"Folha de frequência - {mesAnoSaida}.docx";
                doc.SaveAs(nomeArquivoSaida);

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"\nFolha de ponto preenchida com sucesso!");
                Console.WriteLine($"Salva como '{nomeArquivoSaida}' na pasta do projeto.");
                Console.ResetColor();
            }
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Ocorreu um erro ao processar o documento Word: {ex.Message}");
            Console.ResetColor();
        }
    }
}