using System.Runtime.Versioning;
using System.Text;
using System.Text.RegularExpressions;
using GemBox.Document;
using GemBox.Spreadsheet;
using SlugGenerator;

namespace WordToExcel
{
    [SupportedOSPlatform("windows")]
    internal class Program
    {
        static async Task Main()
        {
            Console.InputEncoding = Console.OutputEncoding = Encoding.Unicode;
            ComponentInfo.SetLicense("TALONEZIO-CRACKED-HEHE");
            var questions = (await ReadQuestionFromWord("Công nghệ điện toán đám mây.doc")).ToList();

            var otherQuestions = ParseQuestionsFromFile("trac-nghiem-cloud.txt");

            questions.AddRange(otherQuestions);

            questions = FilterQuestionsBySlug(questions!)!;

            await ExportToNineQuizExcel(questions, "Cloud Full.xlsx");

            await ExportToNineQuizExcelSplit(questions, "Công nghệ điện toán đám mây.xlsx");

            //await ExportToEduQuizText(questions, "output.txt");

            ImportQuestionsToQuizizz(questions!, "Cloud Quizizz.xlsx");

        }

        private static void ImportQuestionsToQuizizz(IEnumerable<Question> questions, string outputFile)
        {
            var workbook = ExcelFile.Load("QuizizzTemplate.xlsx");

            var worksheet = workbook.Worksheets[0];

            worksheet.Rows.Remove(1, worksheet.Rows.Count - 1);

            var row = 1;

            foreach (var question in questions)
            {


                worksheet.Cells[row, 0].Value = question.Title;

                worksheet.Cells[row, 1].Value = "Multiple Choice";

                worksheet.Cells[row, 2].Value = question.AnswerA;
                worksheet.Cells[row, 3].Value = question.AnswerB;
                worksheet.Cells[row, 4].Value = question.AnswerC;
                worksheet.Cells[row, 5].Value = question.AnswerD;


                worksheet.Cells[row, 7].Value = question.Answer == "A" ? "1" : question.Answer == "B" ? "2" : question.Answer == "C" ? "3" : "4";
                row++;
            }

            workbook.Save(outputFile);
        }

        private static async Task ExportToEduQuizText(List<Question?> questions, string outputTxt)
        {
            var stringBuilder = new StringBuilder();
            foreach (var question in questions.OfType<Question>())
            {
                stringBuilder.AppendLine(question.Title);

                var answers = new List<string> { question.AnswerA, question.AnswerB, question.AnswerC, question.AnswerD };
                var correctAnswer = question.Answer;

                foreach (var answer in answers)
                {
                    if (answer == GetAnswerText(correctAnswer, question))
                    {
                        stringBuilder.AppendLine($"* {answer}");
                    }
                    else
                    {
                        stringBuilder.AppendLine(answer);
                    }
                }

                stringBuilder.AppendLine();  // Dòng trống giữa các câu hỏi
            }

            await File.WriteAllTextAsync(outputTxt, stringBuilder.ToString());
        }
        private static string GetAnswerText(string correctAnswer, Question question)
        {
            return correctAnswer switch
            {
                "A" => question.AnswerA,
                "B" => question.AnswerB,
                "C" => question.AnswerC,
                "D" => question.AnswerD,
                _ => string.Empty
            };
        }

        public static List<Question> FilterQuestionsBySlug(List<Question> questions)
        {
            questions = questions.Select(q =>
                {
                    var answers = new[] { q.AnswerA, q.AnswerB, q.AnswerC, q.AnswerD }
                        .Where(a => !string.IsNullOrEmpty(a))
                        .OrderBy(a => a.Length)
                        .ToArray();

                    var answersSlug = string.Join("-", answers.Select(a => a.GenerateSlug()));
                    var questionSlug = $"{q.Title.GenerateSlug()}-{answersSlug}";

                    return new
                    {
                        Question = q,
                        Slug = questionSlug
                    };
                })
                .DistinctBy(x => x.Slug)
                .Select(x => x.Question)
                .ToList();

            return questions;
        }
        private static async Task ExportToNineQuizExcelSplit(List<Question?> questions, string fileName, int itemPerPage = 100)
        {
            fileName = Path.GetFileNameWithoutExtension(fileName);

            var pageCount = (int)Math.Ceiling((double)questions.Count() / itemPerPage);
            var exportTasks = new List<Task>();

            for (var i = 0; i < pageCount; i++)
            {
                var group = questions.Skip(i * itemPerPage).Take(itemPerPage);
                var groupFileName = $"{fileName}_{i + 1}.xlsx";
                exportTasks.Add(ExportToNineQuizExcel(group, groupFileName));
            }

            await Task.WhenAll(exportTasks);
        }
        private static Task ExportToNineQuizExcel(IEnumerable<Question?> questions, string fileName)
        {
            fileName = Path.ChangeExtension(fileName, ".xlsx");
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

            var filePath = Path.Combine(directory, fileName);
            workbook.Save(filePath);
            Console.WriteLine($"Export hoàn tất: {filePath}");
            return Task.CompletedTask;
        }
        public static List<Question> ParseQuestionsFromFile(string filePath)
        {
            var questions = new List<Question>();
            var lines = File.ReadAllLines(filePath);

            for (int i = 0; i <= lines.Length; i += 6)
            {
                var question = new Question()
                {
                    Title = lines[i],
                    AnswerA = lines[i + 1].TrimStart('*', ' '),
                    AnswerB = lines[i + 2].TrimStart('*', ' '),
                    AnswerC = lines[i + 3].TrimStart('*', ' '),
                    AnswerD = lines[i + 4].TrimStart('*', ' '),
                };
                for (int j = i + 1; j <= i + 4; ++j)
                {
                    if (lines[j].StartsWith("*"))
                    {
                        question.Answer = ((char)(j % 6 - 1 + 65)).ToString();
                        questions.Add(question);
                        break;
                    }
                }
            }

            return questions;
        }
        private static Task<IEnumerable<Question?>> ReadQuestionFromWord(string path)
        {
            var tsc = new TaskCompletionSource<IEnumerable<Question?>>();

            var questions = new List<Question>();

            var document = DocumentModel.Load(path);

            var content = document.Content.ToString();

            string[] questionContents = content.Split(["Câu hỏi"], StringSplitOptions.RemoveEmptyEntries);

            foreach (var questionContent in questionContents)
            {
                var indexOfFirstNewLine = questionContent.IndexOf('\n');

                if (indexOfFirstNewLine == -1) continue;

                string questionInput = questionContent[(indexOfFirstNewLine + 1)..].Trim('\n');

                questionInput = questionInput.Replace("*\r\n", "*")
                    .Replace("\uf0b7", "");

                int indexAnswer = 0;
                var findValue = "\t";
                while (questionInput.IndexOf(findValue, StringComparison.Ordinal) >= 0)
                {
                    int tabIndex = questionInput.IndexOf(findValue, StringComparison.Ordinal);
                    questionInput = questionInput.Remove(tabIndex, 1).Insert(tabIndex, $"{(char)(indexAnswer++ + 'A')}. ");
                }



                var questionSplit = questionInput.Split("\n");

                if (questionSplit.Length < 5)
                    continue;

                // Làm sạch chuỗi đầu vào
                string input = questionSplit[0].Replace(" []:", ":").Trim();

                // Regex để lấy nội dung câu hỏi
                string pattern = @"Câu\s\d+:\s(.+)";
                var matches = Regex.Matches(input, pattern);

                // Lấy tiêu đề dựa trên số lượng match
                string title = matches.Count >= 2
                    ? matches[1].Groups[1].Value // Lấy nội dung từ nhóm bắt thứ nhất
                    : matches.Count == 1
                        ? matches[0].Groups[1].Value // Lấy nội dung nếu chỉ có 1 match
                        : input; // Nếu không có match, dùng toàn bộ chuỗi gốc


                var question = new Question()
                {
                    Title = title,
                    AnswerA = questionSplit[1].Split(' ', 2)[^1].TrimStart('*').Trim().Replace("[<$>] ", "").Replace("a. ", "").Replace("b. ", "").Replace("c. ", "").Replace("d. ", ""),
                    AnswerB = questionSplit[2].Split(' ', 2)[^1].TrimStart('*').Trim().Replace("[<$>] ", "").Replace("a. ", "").Replace("b. ", "").Replace("c. ", "").Replace("d. ", ""),
                    AnswerC = questionSplit[3].Split(' ', 2)[^1].TrimStart('*').Trim().Replace("[<$>] ", "").Replace("a. ", "").Replace("b. ", "").Replace("c. ", "").Replace("d. ", ""),
                    AnswerD = questionSplit[4].Split(' ', 2)[^1].TrimStart('*').Trim().Replace("[<$>] ", "").Replace("a. ", "").Replace("b. ", "").Replace("c. ", "").Replace("d. ", ""),
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
            var pattern = @"(\w)\. \*";

            var match = Regex.Match(answerOptions, pattern);

            return match.Success ? match.Groups[1].Value : string.Empty;
        }
    }
}
