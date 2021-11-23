using DesafioFULL.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace DesafioFULL
{
    class BuscaCEP
    {
        public static Model.Endereco BuscarCEP(string cep)
        {
            Endereco endereco = new Endereco();

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://viacep.com.br/ws/" + cep + "/json/");

            HttpWebResponse servidor = (HttpWebResponse)request.GetResponse();

            if (servidor.StatusCode != HttpStatusCode.OK)
            {
                Console.WriteLine("Server Not Found");
                return null;
            }

            using (Stream webStream = servidor.GetResponseStream())
            {
                if (webStream != null)
                {
                    using (StreamReader responseReader = new StreamReader(webStream))
                    {
                        string response = responseReader.ReadToEnd();

                        response = Regex.Replace(response, "[{},]", string.Empty);

                        response = response.Replace("\"", "");

                        String[] substrings = response.Split('\n');

                        int cont = 0;

                        foreach (var substring in substrings)
                        {
                            if (cont == 1)
                            {
                                string[] valor = substring.Split(":".ToCharArray());

                                endereco.CEP = valor[1];

                                if (valor[0] == "  erro")
                                {
                                    return null;
                                }
                            }

                            if (cont == 2)
                            {
                                string[] valor = substring.Split(":".ToCharArray());
                                endereco.Logradouro = valor[1];
                            }

                            if (cont == 4)
                            {
                                string[] valor = substring.Split(":".ToCharArray());
                                endereco.Bairro = valor[1];
                            }

                            if (cont == 5)
                            {
                                string[] valor = substring.Split(":".ToCharArray());
                                endereco.Cidade = valor[1];
                            }

                            if (cont == 6)
                            {
                                string[] valor = substring.Split(":".ToCharArray());
                                endereco.UF = valor[1];
                            }

                            cont++;
                        }
                    }
                }
                return endereco;
            }
        }

        public static void GerarResultado(List<Endereco> enderecos)
        {
            Console.WriteLine("Criando o arquivo...");

            var result = new XLWorkbook();

            var planilha = result.Worksheets.Add("Resultado");

            planilha.Cell("A1").Value = "CEP";
            planilha.Cell("B1").Value = "Logradouro";
            planilha.Cell("C1").Value = "Bairro/Distrito";
            planilha.Cell("D1").Value = "Localidade/UF";
            planilha.Cell("E1").Value = "Data/Hora processamento";

            var linha = 2;

            foreach (Endereco x in enderecos)
            {
                planilha.Cell("A" + linha.ToString()).Value = x.CEP;
                planilha.Cell("B" + linha.ToString()).Value = x.Logradouro;
                planilha.Cell("C" + linha.ToString()).Value = x.Bairro;
                planilha.Cell("D" + linha.ToString()).Value = x.Cidade + "/" + x.UF;
                planilha.Cell("E" + linha.ToString()).Value = DateTime.Now;
                linha++;
            }

            result.Dispose();

            Console.ReadKey();
        }
    }
}
