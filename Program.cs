using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using DesafioFULL.Model;

namespace DesafioFULL
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Endereco> enderecos = new List<Endereco>();
            
            Console.WriteLine(" Buscando informações... \n");

            var linha = 2;

            var listarCEP = new XLWorkbook("../Arquivo/Lista_de_CEPs.xlsx");
            var result = listarCEP.Worksheet(1);

            
            while (true)
            {
                var faixaCEP = result.Cell("A" + linha.ToString()).Value.ToString();
                var iniCEP = result.Cell("B" + linha.ToString()).Value.ToString();
                var fimCEP = result.Cell("C" + linha.ToString()).Value.ToString();

                if (string.IsNullOrEmpty(faixaCEP)) break;

                for (int i = Convert.ToInt32(iniCEP); i <= Convert.ToInt32(fimCEP); i++)
                {
                    Endereco encontrouEnd = BuscaCEP.BuscarCEP(i.ToString());
                    if (encontrouEnd != null)
                        enderecos.Add(encontrouEnd);
                }
                linha++;
            }
            BuscaCEP.GerarResultado(enderecos);
        }
    }
}
