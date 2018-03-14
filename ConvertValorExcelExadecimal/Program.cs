using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConvertValorExcelHexadecimal
{
    class Program
    {
        static void Main(string[] args)
        {
            //método que abre o excel
            openExcel();
        }

        private static void openExcel()
        {
            //indique o endereço o arquivo xlsx
            var path = "C:\\Users\\MauricioJunior\\Documents\\Visual Studio 2015\\Projects\\ConvertValorExcelExadecimal\\ConvertValorExcelExadecimal\\bin\\Debug\\2fevereiro.xlsx";

            FileInfo fi = new FileInfo(path);
            if (fi.Exists)
            {
                //Excel - nome utilizado no using (início da classe)
                //Abre o excel
                Excel._Application oApp = new Excel.Application();
                oApp.Visible = false; //Não mostra para o usuário que abriu o arquivo

                //Abre
                Excel.Workbook oWorkbook = oApp.Workbooks.Open(path);

                //Abre o Worksheet nome [Fev 01]
                Excel.Worksheet oWorksheet = oWorkbook.Worksheets["Fev 01"];

                //Navega nas linhas e colunas
                NavegarColunasELinhas(oWorksheet);

                //oWorkbook.Save();

                //fecha o workbook
                oWorkbook.Close();

                //fecha o arquivo
                oApp.Quit();

                //seta para null para retirar da memória
                oWorksheet = null;
                oWorkbook = null;
                oApp = null;

                //retira do Garbage Collection
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            else
            {
                //o arquivo não existe
                Console.WriteLine("O arquivo não existe");
            }
        }

        private static void NavegarColunasELinhas(Excel.Worksheet oWorksheet)
        {
            //Pega a quantidade de coluna e linha
            int numeroColuna = oWorksheet.UsedRange.Columns.Count;
            int numeroLinha = oWorksheet.UsedRange.Rows.Count;

            //Lê o arquivo linha por linha através do array
            object[,] array = oWorksheet.UsedRange.Value;
            for (int j = 1; j <= numeroColuna; j++)
            {
                for (int i = 1; i <= numeroLinha; i++)
                {
                    if (array[i, j] != null)
                        //if (array[i, j].ToString() == "1900006765") //Verifica a coluna C se o valor é igual
                        //{
                        for (int m = i + 1; m < numeroLinha; m++)
                        {
                            //voce pega [linha, coluna] - depende da maneira que está na sua planilha
                            //if (Convert.ToInt32(array[m, j].ToString()) > 50)
                            //{
                            //array[m, j + 1] = "Yes";
                            //}
                        }

                    //coloca o valor da coluna de volta.
                    oWorksheet.UsedRange.Value = array; //dados no array
                    return;
                    //}
                }
            }

        }
    }
}
