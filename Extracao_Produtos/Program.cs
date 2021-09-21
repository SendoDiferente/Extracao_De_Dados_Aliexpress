using Extracao_Produtos.Site;
using System;

namespace Extracao_Produtos
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Voce quer que seja feito uma atualizacao dos dados ");
            Console.WriteLine("1- Sim 2- Nao");
            string resposta = Console.ReadLine();
            Console.WriteLine("Voce quer que seja extracao em massa");
            Console.WriteLine("1- Sim 2- Nao");
            string resposta1 = Console.ReadLine();
            Sites start = new Sites();
            start.Aliexpress(resposta, resposta1);
        }
    }
}