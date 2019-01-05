using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Globalization;

namespace ExtetecaoWord
{
    class Braile
    {
        // Dictionary<char, bool[][]> dict = new Dictionary<int, string>()
        static private IDictionary<char, byte[,]> dict = new Dictionary<char, byte[,]>()
                                           {
                                        {'a',
                                         new byte[3,2]
                                            {{1,0},
                                             {0,0},
                                             {0,0}}
                                        },
                                        {'b',
                                         new byte[3,2]
                                            {{1,0},
                                             {1,0},
                                             {0,0}}
                                        },
                                        {'c',
                                         new byte[3,2]
                                            {{1,1},
                                             {0,0},
                                             {0,0}}
                                        },
                                        {'d',
                                         new byte[3,2]
                                            {{1,1},
                                             {0,1},
                                             {0,0}}
                                        },
                                        {'e',
                                         new byte[3,2]
                                            {{1,0},
                                             {0,1},
                                             {0,0}}
                                        },
                                        {'f',
                                         new byte[3,2]
                                            {{1,1},
                                             {1,0},
                                             {0,0}}
                                        },
                                        {'g',
                                         new byte[3,2]
                                            {{1,1},
                                             {1,1},
                                             {0,0}}
                                        },
                                        {'h',
                                         new byte[3,2]
                                            {{1,0},
                                             {1,1},
                                             {0,0}}
                                        },
                                        {'i',
                                         new byte[3,2]
                                            {{0,1},
                                             {1,0},
                                             {0,0}}
                                        },
                                        {'j',
                                         new byte[3,2]
                                            {{0,1},
                                             {1,1},
                                             {0,0}}
                                        },
                                        {'k',
                                         new byte[3,2]
                                            {{1,0},
                                             {0,0},
                                             {1,0}}
                                        },
                                        {'l',
                                         new byte[3,2]
                                            {{1,0},
                                             {1,0},
                                             {1,0}}
                                        },
                                        {'m',
                                         new byte[3,2]
                                            {{1,1},
                                             {0,0},
                                             {1,0}}
                                        },
                                        {'n',
                                         new byte[3,2]
                                            {{1,1},
                                             {0,1},
                                             {1,0}}
                                        },
                                        {'o',
                                         new byte[3,2]
                                            {{1,0},
                                             {0,1},
                                             {1,0}}
                                        },
                                        {'p',
                                         new byte[3,2]
                                            {{1,1},
                                             {1,0},
                                             {1,0}}
                                        },
                                        {'q',
                                         new byte[3,2]
                                            {{1,1},
                                             {1,1},
                                             {1,0}}
                                        },
                                        {'r',
                                         new byte[3,2]
                                            {{1,0},
                                             {1,1},
                                             {1,0}}
                                        },
                                        {'s',
                                         new byte[3,2]
                                            {{0,1},
                                             {1,0},
                                             {1,0}}
                                        },
                                        {'t',
                                         new byte[3,2]
                                            {{0,1},
                                             {1,1},
                                             {1,0}}
                                        },
                                        {'u',
                                         new byte[3,2]
                                            {{1,0},
                                             {0,0},
                                             {1,1}}
                                        },
                                        {'v',
                                         new byte[3,2]
                                            {{1,0},
                                             {1,0},
                                             {1,1}}
                                        },
                                        {'w',
                                         new byte[3,2]
                                            {{0,1},
                                             {1,1},
                                             {0,1}}
                                        },
                                        {'x',
                                         new byte[3,2]
                                            {{1,1},
                                             {0,0},
                                             {1,1}}
                                        },
                                        {'y',
                                         new byte[3,2]
                                            {{1,1},
                                             {0,1},
                                             {1,1}}
                                        },
                                        {'z',
                                         new byte[3,2]
                                            {{1,0},
                                             {0,1},
                                             {1,1}}
                                        },
                                        {'ç',
                                         new byte[3,2]
                                            {{1,1},
                                             {1,0},
                                             {1,1}}
                                        },
                                        {'á',
                                         new byte[3,2]
                                            {{1,0},
                                             {1,1},
                                             {1,1}}
                                        },
                                        {'é',
                                         new byte[3,2]
                                            {{1,1},
                                             {1,1},
                                             {1,1}}
                                        },
                                        {'í',
                                         new byte[3,2]
                                            {{0,1},
                                             {0,0},
                                             {1,0}}
                                        },
                                        {'ó',
                                         new byte[3,2]
                                            {{0,1},
                                             {0,0},
                                             {1,1}}
                                        },
                                        {'ú',
                                         new byte[3,2]
                                            {{0,1},
                                             {1,1},
                                             {1,1}}
                                        },
                                        {'â',
                                         new byte[3,2]
                                            {{1,0},
                                             {0,0},
                                             {0,1}}
                                        },
                                        {'ê',
                                         new byte[3,2]
                                            {{1,0},
                                             {1,0},
                                             {0,1}}
                                        },
                                        {'ô',
                                         new byte[3,2]
                                            {{1,1},
                                             {0,1},
                                             {0,1}}
                                        },
                                        {'ã',
                                         new byte[3,2]
                                            {{0,1},
                                             {0,1},
                                             {1,0}}
                                        },
                                        {'õ',
                                         new byte[3,2]
                                            {{0,1},
                                             {1,0},
                                             {0,1}}
                                        },
                                        // Pontuação
                                        {'.',
                                         new byte[3,2]
                                            {{0,0},
                                             {0,0},
                                             {1,0}}
                                        },
                                        {',',
                                         new byte[3,2]
                                            {{0,0},
                                             {1,0},
                                             {0,0}}
                                        },
                                        {'?',
                                         new byte[3,2]
                                            {{0,0},
                                             {1,0},
                                             {0,1}}
                                        },
                                        {';',
                                         new byte[3,2]
                                            {{0,0},
                                             {1,0},
                                             {1,0}}
                                        },
                                        {':',
                                         new byte[3,2]
                                            {{0,0},
                                             {1,1},
                                             {0,0}}
                                        },
                                        {'!',
                                         new byte[3,2]
                                            {{0,0},
                                             {1,1},
                                             {1,0}}
                                        },
                                        {'-',
                                         new byte[3,2]
                                            {{0,0},
                                             {0,0},
                                             {1,1}}
                                        },
                                        {' ',
                                         new byte[3,2]
                                            {{0,0},
                                             {0,0},
                                             {0,0}}
                                        },
                                        {'\'',/* verificar! */ 
                                         new byte[3,2]
                                            {{0,0},
                                             {1,1},
                                             {0,0}}
                                        },
                                        {'"',
                                         new byte[3,2]
                                            {{0,0},
                                             {1,0},
                                             {1,1}}
                                        },
                                        {'/',
                                         new byte[3,2]
                                            {{0,1},
                                             {0,0},
                                             {1,0}}
                                        },
                                        {'$',
                                         new byte[3,2]
                                            {{0,0},
                                             {0,1},
                                             {0,1}}
                                        },
                                        {'@',
                                         new byte[3,2]
                                            {{1,0},
                                             {0,1},
                                             {0,1}}
                                        },
                                        {'=',
                                         new byte[3,2]
                                            {{0,0},
                                             {1,1},
                                             {1,1}}
                                        },
                                        {'^',
                                         new byte[3,2]
                                            {{1,0},
                                             {0,0},
                                             {0,1}}
                                        },
                                        {'°',
                                         new byte[3,2]
                                            {{0,0},
                                             {0,1},
                                             {1,1}}
                                        },
                                        // Numeros
                                        {'1',
                                         new byte[3, 2]
                                            {{1,0},
                                             {0,0},
                                             {0,0}}
                                        },
                                        {'2',
                                         new byte[3, 2]
                                            {{1,0},
                                             {1,0},
                                             {0,0}}
                                        },
                                        {'3',
                                         new byte[3, 2]
                                            {{1,1},
                                             {0,0},
                                             {0,0}}
                                        },
                                        {'4',
                                         new byte[3, 2]
                                            {{1,1},
                                             {0,1},
                                             {0,0}}
                                        },
                                        {'5',
                                         new byte[3, 2]
                                            {{1,0},
                                             {0,1},
                                             {0,0}}
                                        },
                                        {'6',
                                         new byte[3, 2]
                                            {{1,1},
                                             {1,0},
                                             {0,0}}
                                        },
                                        {'7',
                                         new byte[3, 2]
                                            {{1,1},
                                             {1,1},
                                             {0,0}}
                                        },
                                        {'8',
                                         new byte[3, 2]
                                            {{1,0},
                                             {1,1},
                                             {0,0}}
                                        },
                                        {'9',
                                         new byte[3, 2]
                                            {{0,1},
                                             {1,0},
                                             {0,0}}
                                        },
                                        {'0',
                                         new byte[3, 2]
                                            {{0,1},
                                             {1,1},
                                             {0,0}}
                                        },
                                        // Especial Braile
                                        {'Ʊ', // Maiuscula
                                         new byte[3, 2]
                                            {{0,1},
                                             {0,0},
                                             {0,1}}
                                        },
                                        {'Ʃ', //Numero
                                         new byte[3, 2]
                                            {{0,1},
                                             {0,0},
                                             {0,1}}
                                        }
                                    };

        private string text = "";

        public Braile()
        {
            Word.Paragraphs paragraphs = Globals.ThisAddIn.Application.ActiveDocument.Paragraphs;

            foreach (Word.Paragraph paragraph in paragraphs)
            {
                this.text += paragraph.Range.Text + "\n";
            }
        }

        public string getBruto()
        {
            return this.text;
        }

        public string getPseudoBraile()
        {
            string[] arrayText = this.text.Split(' ');

            for (int i = 0; i < arrayText.Length; i++)
            {
                if (Regex.IsMatch(arrayText[i], @"\d"))
                {
                    arrayText[i] = 'Ʃ' + arrayText[i];
                    continue;
                }

                CultureInfo cultureInfo = CultureInfo.CurrentCulture;
                TextInfo textInfo = cultureInfo.TextInfo;

                if (arrayText[i] == textInfo.ToUpper(arrayText[i]))
                {
                    arrayText[i] = "ƱƱ" + arrayText[i].ToLower();
                    continue;
                }
                if (arrayText[i] == textInfo.ToTitleCase(arrayText[i]))
                {
                    arrayText[i] = 'Ʊ' + arrayText[i].ToLower();
                    continue;
                }
            }
            return String.Join(" ", arrayText);
        }

        public byte[][,] getBraile()
        {
            List<byte[,]> list = new List<byte[,]>();

            foreach (char c in this.getPseudoBraile())
            {
                if(Braile.dict.ContainsKey(c))
                {
                    list.Add(Braile.dict[c]);
                }
                else
                {
                    MessageBox.Show(c + " Não encontrado no Dicionario de conversão!");
                }
            }
            return list.ToArray();
        }

        public List<byte[][,]> getPage()
        {
            const byte perCollun = 5;

            byte[][,] list = this.getBraile();
            List<byte[][,]> page = new List<byte[][,]>();

            for (int i = 0; i < list.Length % perCollun; i++)
            {
                page.Add(list.Skip(i * perCollun).Take((i + 1) * perCollun).ToArray());
            }
            // 


            return page;
        }

        public byte[][] getLineSerializationToArduino()
        {
            List<byte[][,]> page = this.getPage();
            List<byte[]> lineSerial = new List<byte[]>();

            for (int i = 0; i < page.Count; i++) //Linha da Pagina
            {
                for (int line = 0; line < 3; line++)// Linha da Celula
                {
                    List<byte> acc = new List<byte>();
                    for (int j = 0; j < page[i].Length; j++) //Celula Da pagina
                    {
                        for (int col = 0; col < 2; col++)
                        {
                            acc.Add(page[i][j][line, col]);
                        }
                    }
                    lineSerial.Add(acc.ToArray());
                }
            }

            return lineSerial.ToArray();
        }

        public byte[] getSerializationToArduino()
        {
            byte[][] lines = this.getLineSerializationToArduino();
            List<byte> serializate = new List<byte>();

            for (int i = 0; i < lines.Length; i++)
            {
                if(i%2==0) // se for par le direto
                {
                    for (int j = 0; j < lines[i].Length; j++)
                    {
                        serializate.Add(lines[i][j]);
                    }
                }
                else //se for impar le inverso
                {
                    for (int j = lines[i].Length-1; j >= 0; j--)
                    {
                        serializate.Add(lines[i][j]);
                    }
                }
            }

            return serializate.ToArray();
        }
    }
}
