using System.Runtime.Versioning;
using System.Text.RegularExpressions;
using GemBox.Document;
using GemBox.Spreadsheet;

namespace WordToExcel
{
    [SupportedOSPlatform("windows")]
    internal class Program
    {
        static async Task Main()
        {
            ComponentInfo.SetLicense("TALONEZIO-CRACKED-HEHE");
            var questions = await ReadQuestionFromWord("Chương trình dịch.doc");
            await ExportToExcel(questions,"Chương trình dịch.xlsx");
        }
        private static Task ExportToExcel(IEnumerable<Question> questions, string fileName)
        {
            SpreadsheetInfo.SetLicense("TalonEzio-Cracked-Hehe");

            var workbook = new ExcelFile();

            var worksheet = workbook.Worksheets.Add("Sheet1");

            worksheet.Cells["A1"].Value = "Câu hỏi";
            worksheet.Cells["B1"].Value = "Đáp án";
            worksheet.Cells["C1"].Value = "Câu trả lời A";
            worksheet.Cells["D1"].Value = "Câu trả lời B";
            worksheet.Cells["E1"].Value = "Câu trả lời C";
            worksheet.Cells["F1"].Value = "Câu trả lời D";

            var row = 1;
            foreach (var question in questions)
            {
                worksheet.Cells[row, 0].Value = question.Title;
                worksheet.Cells[row, 1].Value = question.Answer;
                worksheet.Cells[row, 2].Value = question.AnswerA;
                worksheet.Cells[row, 3].Value = question.AnswerB;
                worksheet.Cells[row, 4].Value = question.AnswerC;
                worksheet.Cells[row, 5].Value = question.AnswerD;
                row++;
            }

            var directory = Path.Combine(Environment.CurrentDirectory, "Exports");
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            string filePath = Path.Combine(directory, fileName);
            workbook.Save(filePath);
            Console.WriteLine($"Export hoàn tất: {filePath}");
            return Task.CompletedTask;
        }
        private static Task<IEnumerable<Question>> ReadQuestionFromWord(string path)
        {
            var tsc = new TaskCompletionSource<IEnumerable<Question>>();

            var questions = new List<Question>();

            var document = DocumentModel.Load(path);

            var content = document.Content.ToString();

            content = content.Replace("\t*\r\n", "*").Replace("\t","").Replace("\uf0b7","");
            string[] questionContents = content.Split(["Câu hỏi"], StringSplitOptions.RemoveEmptyEntries);

            foreach (var questionContent in questionContents)
            {
                var indexOfFirstNewLine = questionContent.IndexOf('\n');

                if (indexOfFirstNewLine == -1) continue;

                string questionInput = questionContent[(indexOfFirstNewLine + 1)..].Trim('\n');

                var questionSplit = questionInput.Split("\n");

                //chỉ tìm câu hỏi có 4 câu trả lời, câu có 3 câu trả lời bỏ qua
                if (questionSplit.Length < 5) 
                    continue;

                var question = new Question()
                {
                    Title = questionSplit[0].Split(" : ", StringSplitOptions.RemoveEmptyEntries)[^1],
                    AnswerA = questionSplit[1].TrimStart('*').Split(' ',2)[^1].Trim(),
                    AnswerB = questionSplit[2].TrimStart('*').Split(' ', 2)[^1].Trim(),
                    AnswerC = questionSplit[3].TrimStart('*').Split(' ', 2)[^1].Trim(),
                    AnswerD = questionSplit[4].TrimStart('*').Split(' ', 2)[^1].Trim(),
                    Answer = FindCorrectAnswer(questionInput)
                };

                questions.Add(question);

            }

            if (questions.Any())
                tsc.SetResult(questions);
            else
            {
                tsc.SetException([new Exception("Không đọc được câu hỏi nào")]);
            }

            return tsc.Task;
        }
        private static string FindCorrectAnswer(string answerOptions)
        {
            var pattern = @"\*(\w)\.";

            var match = Regex.Match(answerOptions, pattern);

            return match.Success ? match.Groups[1].Value : string.Empty;
        }
    }
}
