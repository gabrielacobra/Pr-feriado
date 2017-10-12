using System;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace arquivocsv
{
    class Program
    {
        static void Main(string[] args)
        {
            Application ex = new Application();
            ex.Visible = true;
            ex.Workbooks.Add();

            /* 
            string nome,email;
            int idade;

            Console.WriteLine("Digite o seu nome: ");
            nome = Console.ReadLine();

            Console.WriteLine("Digite o seu e-mail: ");
            email = Console.ReadLine();

            Console.WriteLine("Digite a sua idade: ");
            idade = Int16.Parse(Console.ReadLine());
            */

            // FileInfo fi = new FileInfo("dados_cabecalho.csv");
                     
            // StreamWriter arquivo;
            
            // if(fi.Exists) {
            // arquivo = new StreamWriter("dados_cabecalho.csv",true);
            // arquivo.WriteLine(nome+";"+email+";"+idade+";"+DateTime.Now.ToShortDateString()); //se ele existe, não é necessário criar o cabecalho
            // }
            // else    {
            // arquivo = new StreamWriter("dados_cabecalho.csv",true); //new StreamWriter gera um objeto.
            // arquivo.WriteLine("Nome;Email;Idade;Data de Cadastro"); //se o arquivo não existe, o comando é para criar o cabecalho
            // arquivo.WriteLine(nome+";"+email+";"+idade+";"+DateTime.Now.ToShortDateString());
            // }
                        
            // arquivo.Close();


            
        }
    }
}
