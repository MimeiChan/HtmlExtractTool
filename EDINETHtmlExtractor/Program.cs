using System;
using System.IO;
using System.Text;
using HtmlAgilityPack;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace EDINETHtmlExtractor
{
    /// <summary>
    /// EDINETの有価証券報告書からHTMLセクションを抽出するクラス
    /// </summary>
    public class HtmlSectionExtractor
    {
        // 抽出対象の開始見出し
        private const string START_SECTION_TITLE = "【損益計算書】";
        // 抽出対象の終了見出し
        private const string END_SECTION_TITLE = "【資本変動計算書】";

        // 抽出状態を管理する変数
        private bool extractionStarted = false;
        private bool extractionCompleted = false;
        private List<string> extractedContentList = new List<string>();
        
        // スタイル情報を保持する変数（新規追加）
        private List<string> extractedStylesList = new List<string>();

        /// <summary>
        /// ディレクトリ内のHTMLファイルから指定されたセクションを抽出する
        /// </summary>
        /// <param name="directoryPath">HTMLファイルが存在するディレクトリのパス</param>
        /// <param name="outputFilePath">出力ファイルのパス</param>
        /// <returns>抽出に成功したかどうか</returns>
        public bool ExtractAndSaveSectionFromDirectory(string directoryPath, string outputFilePath)
        {
            try
            {
                // ディレクトリの存在確認
                if (!Directory.Exists(directoryPath))
                {
                    Console.WriteLine($"ディレクトリが見つかりません: {directoryPath}");
                    return false;
                }

                // ディレクトリ内のHTMLファイルを取得して連番順にソート
                List<string> htmlFiles = GetSortedHtmlFiles(directoryPath);
                if (htmlFiles.Count == 0)
                {
                    Console.WriteLine("指定されたディレクトリにHTMLファイルが見つかりませんでした。");
                    return false;
                }

                // 状態をリセット
                extractionStarted = false;
                extractionCompleted = false;
                extractedContentList.Clear();
                extractedStylesList.Clear(); // スタイルリストもクリア

                // 各ファイルを順番に処理
                foreach (string htmlFile in htmlFiles)
                {
                    Console.WriteLine($"処理中: {htmlFile}");
                    ProcessHtmlFile(htmlFile);

                    // 抽出が完了したら終了
                    if (extractionCompleted)
                        break;
                }

                // 抽出が始まったが完了しなかった場合
                if (extractionStarted && !extractionCompleted)
                {
                    Console.WriteLine("警告: 【損益計算書】セクションは見つかりましたが、【資本変動計算書】セクションが見つかりませんでした。");
                    Console.WriteLine("最後のファイルまで抽出を行います。");
                }

                // 抽出したHTMLを保存
                if (extractedContentList.Count > 0)
                {
                    SaveToFile(extractedContentList, extractedStylesList, outputFilePath);
                    Console.WriteLine($"セクションの抽出に成功しました。出力ファイル: {outputFilePath}");
                    return true;
                }
                else
                {
                    Console.WriteLine("【損益計算書】セクションが見つかりませんでした。");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"エラーが発生しました: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 単一のHTMLファイルを処理する
        /// </summary>
        /// <param name="htmlFilePath">処理するHTMLファイルのパス</param>
        private void ProcessHtmlFile(string htmlFilePath)
        {
            try
            {
                // HTMLファイルを読み込む
                HtmlDocument doc = new HtmlDocument();
                doc.Load(htmlFilePath, Encoding.UTF8);

                // スタイル情報を抽出（新規追加）
                ExtractStyles(doc);

                // 可能性のある見出し要素を定義
                string[] headingTags = { "h1", "h2", "h3", "h4", "h5", "h6", "div", "span", "p" };

                // 開始ノードと終了ノードを検索
                HtmlNode startNode = null;
                HtmlNode endNode = null;

                // 見出しの検索 - 修正：常に終了見出しを検索するように変更
                foreach (var tag in headingTags)
                {
                    var nodes = doc.DocumentNode.SelectNodes($"//{tag}");
                    if (nodes == null) continue;

                    // 開始見出しの検索（まだ抽出が始まっていない場合）
                    if (!extractionStarted && startNode == null)
                    {
                        startNode = nodes.FirstOrDefault(n => n.InnerText.Contains(START_SECTION_TITLE));
                    }

                    // 終了見出しの検索（常に検索する）
                    if (endNode == null)
                    {
                        endNode = nodes.FirstOrDefault(n => n.InnerText.Contains(END_SECTION_TITLE));
                    }
                }

                // 同一ファイル内に開始見出しと終了見出しの両方がある場合
                if (startNode != null && endNode != null && !extractionStarted)
                {
                    Console.WriteLine("同一ファイル内に開始見出しと終了見出しの両方が見つかりました。");
                    extractionStarted = true;
                    ExtractContentBetweenNodes(startNode, endNode, doc);
                    extractionCompleted = true;  // このファイルで抽出完了とマーク
                }
                // ファイル内に開始見出しのみがある場合
                else if (startNode != null && !extractionStarted)
                {
                    Console.WriteLine("開始見出しが見つかりました。");
                    extractionStarted = true;
                    ExtractContentFromStartNode(startNode, doc);
                }
                // 既に抽出が始まっていて、このファイルに終了見出しがある場合
                else if (extractionStarted && endNode != null && !extractionCompleted)
                {
                    Console.WriteLine("終了見出しが見つかりました。");
                    ExtractContentUntilEndNode(endNode, doc);
                    extractionCompleted = true;
                }
                // 既に抽出が始まっていて、このファイルに終了見出しがない場合
                else if (extractionStarted && !extractionCompleted)
                {
                    Console.WriteLine("抽出中：このファイルは全体を抽出します。");
                    // ファイル全体を抽出
                    ExtractEntireFile(doc);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ファイル処理中にエラーが発生しました: {htmlFilePath}, {ex.Message}");
            }
        }

        /// <summary>
        /// HTMLドキュメントからスタイル情報を抽出する（新規追加）
        /// </summary>
        /// <param name="doc">HTMLドキュメント</param>
        private void ExtractStyles(HtmlDocument doc)
        {
            // headタグ内のstyleタグを抽出
            var styleNodes = doc.DocumentNode.SelectNodes("//head/style");
            if (styleNodes != null)
            {
                foreach (var styleNode in styleNodes)
                {
                    if (!string.IsNullOrWhiteSpace(styleNode.InnerHtml) && 
                        !extractedStylesList.Contains(styleNode.InnerHtml))
                    {
                        extractedStylesList.Add(styleNode.InnerHtml);
                        Console.WriteLine("スタイル情報を抽出しました。");
                    }
                }
            }

            // headタグ内のlinkタグ（外部スタイルシート）を抽出
            var linkNodes = doc.DocumentNode.SelectNodes("//head/link[@rel='stylesheet']");
            if (linkNodes != null)
            {
                foreach (var linkNode in linkNodes)
                {
                    string href = linkNode.GetAttributeValue("href", "");
                    if (!string.IsNullOrWhiteSpace(href) && !extractedStylesList.Contains(linkNode.OuterHtml))
                    {
                        extractedStylesList.Add(linkNode.OuterHtml);
                        Console.WriteLine("外部スタイルシートの参照を抽出しました。");
                    }
                }
            }
        }

        /// <summary>
        /// 開始ノードと終了ノードの間のコンテンツを抽出する
        /// </summary>
        /// <param name="startNode">開始ノード</param>
        /// <param name="endNode">終了ノード</param>
        /// <param name="doc">HtmlDocument</param>
        private void ExtractContentBetweenNodes(HtmlNode startNode, HtmlNode endNode, HtmlDocument doc)
        {
            // 修正：テーブル構造とスタイルを保持するために、DOMツリーから直接抽出する方法に変更
            HtmlDocument extractedDoc = new HtmlDocument();
            HtmlNode bodyNode = extractedDoc.CreateElement("div");
            extractedDoc.DocumentNode.AppendChild(bodyNode);

            bool isCapturing = false;
            
            // 親ノードを見つける関数
            HtmlNode FindCommonAncestor(HtmlNode node1, HtmlNode node2)
            {
                var ancestors1 = new List<HtmlNode>();
                var current = node1;
                while (current != null)
                {
                    ancestors1.Add(current);
                    current = current.ParentNode;
                }

                current = node2;
                while (current != null)
                {
                    if (ancestors1.Contains(current))
                    {
                        return current;
                    }
                    current = current.ParentNode;
                }

                return doc.DocumentNode.SelectSingleNode("//body");
            }

            // 共通の親ノードを見つける
            HtmlNode commonAncestor = FindCommonAncestor(startNode, endNode);
            Console.WriteLine($"共通の親ノード: {commonAncestor.Name}");

            // 開始ノードから終了ノードまでを抽出する
            void ProcessNode(HtmlNode node, HtmlNode parentInNewDoc)
            {
                // 開始ノードに到達
                if (node == startNode)
                {
                    isCapturing = true;
                }

                // 終了ノードに到達
                if (node == endNode)
                {
                    isCapturing = false;
                    return;
                }

                if (isCapturing)
                {
                    // ノードをコピーして新しいドキュメントに追加
                    HtmlNode newNode = extractedDoc.CreateElement(node.Name);
                    
                    // 属性をコピー（スタイル情報を保持）
                    foreach (var attr in node.Attributes)
                    {
                        newNode.SetAttributeValue(attr.Name, attr.Value);
                    }

                    // ノードコンテンツの処理
                    if (!node.HasChildNodes)
                    {
                        newNode.InnerHtml = node.InnerHtml;
                    }

                    parentInNewDoc.AppendChild(newNode);

                    // 子ノードを再帰的に処理
                    foreach (var child in node.ChildNodes)
                    {
                        ProcessNode(child, newNode);
                    }
                }
                else if (!isCapturing && node != endNode)
                {
                    // 子ノードに対して再帰（開始ノードを探すため）
                    foreach (var child in node.ChildNodes)
                    {
                        ProcessNode(child, parentInNewDoc);
                    }
                }
            }

            // 処理開始
            ProcessNode(commonAncestor, bodyNode);

            // 抽出されたコンテンツを追加
            extractedContentList.Add(bodyNode.InnerHtml);
            Console.WriteLine($"抽出されたコンテンツのサイズ: {bodyNode.InnerHtml.Length} 文字");
        }

        /// <summary>
        /// 開始ノードから文書の最後までのコンテンツを抽出する
        /// </summary>
        /// <param name="startNode">開始ノード</param>
        /// <param name="doc">HtmlDocument</param>
        private void ExtractContentFromStartNode(HtmlNode startNode, HtmlDocument doc)
        {
            // 修正：テーブル構造とスタイルを保持するために、DOMツリーから直接抽出する方法に変更
            HtmlDocument extractedDoc = new HtmlDocument();
            HtmlNode bodyNode = extractedDoc.CreateElement("div");
            extractedDoc.DocumentNode.AppendChild(bodyNode);

            bool isCapturing = false;

            // 階層順でコンテンツを収集
            void ProcessNode(HtmlNode node, HtmlNode parentInNewDoc)
            {
                // 開始ノードに到達
                if (node == startNode)
                {
                    isCapturing = true;
                }

                if (isCapturing)
                {
                    // ノードをコピーして新しいドキュメントに追加
                    HtmlNode newNode = extractedDoc.CreateElement(node.Name);
                    
                    // 属性をコピー（スタイル情報を保持）
                    foreach (var attr in node.Attributes)
                    {
                        newNode.SetAttributeValue(attr.Name, attr.Value);
                    }

                    // ノードコンテンツの処理
                    if (!node.HasChildNodes)
                    {
                        newNode.InnerHtml = node.InnerHtml;
                    }

                    parentInNewDoc.AppendChild(newNode);

                    // 子ノードを再帰的に処理
                    foreach (var child in node.ChildNodes)
                    {
                        ProcessNode(child, newNode);
                    }
                }
                else
                {
                    // 子ノードに対して再帰（開始ノードを探すため）
                    foreach (var child in node.ChildNodes)
                    {
                        ProcessNode(child, parentInNewDoc);
                    }
                }
            }

            // BODYタグ以下の要素を対象に
            HtmlNode docBodyNode = doc.DocumentNode.SelectSingleNode("//body");
            if (docBodyNode == null)
            {
                Console.WriteLine("警告: bodyタグが見つかりませんでした。");
                return;
            }

            // 処理開始
            ProcessNode(docBodyNode, bodyNode);

            // 抽出されたコンテンツを追加
            extractedContentList.Add(bodyNode.InnerHtml);
            Console.WriteLine($"抽出されたコンテンツのサイズ: {bodyNode.InnerHtml.Length} 文字");
        }

        /// <summary>
        /// 文書の先頭から終了ノードまでのコンテンツを抽出する
        /// </summary>
        /// <param name="endNode">終了ノード</param>
        /// <param name="doc">HtmlDocument</param>
        private void ExtractContentUntilEndNode(HtmlNode endNode, HtmlDocument doc)
        {
            // 修正：テーブル構造とスタイルを保持するために、DOMツリーから直接抽出する方法に変更
            HtmlDocument extractedDoc = new HtmlDocument();
            HtmlNode bodyNode = extractedDoc.CreateElement("div");
            extractedDoc.DocumentNode.AppendChild(bodyNode);

            bool isCapturing = true;

            // 階層順でコンテンツを収集
            void ProcessNode(HtmlNode node, HtmlNode parentInNewDoc)
            {
                // 終了ノードに到達
                if (node == endNode)
                {
                    isCapturing = false;
                    return;
                }

                if (isCapturing)
                {
                    // ノードをコピーして新しいドキュメントに追加
                    HtmlNode newNode = extractedDoc.CreateElement(node.Name);
                    
                    // 属性をコピー（スタイル情報を保持）
                    foreach (var attr in node.Attributes)
                    {
                        newNode.SetAttributeValue(attr.Name, attr.Value);
                    }

                    // ノードコンテンツの処理
                    if (!node.HasChildNodes)
                    {
                        newNode.InnerHtml = node.InnerHtml;
                    }

                    parentInNewDoc.AppendChild(newNode);

                    // 子ノードを再帰的に処理
                    foreach (var child in node.ChildNodes)
                    {
                        ProcessNode(child, newNode);
                    }
                }
            }

            // BODYタグ以下の要素を対象に
            HtmlNode docBodyNode = doc.DocumentNode.SelectSingleNode("//body");
            if (docBodyNode == null)
            {
                Console.WriteLine("警告: bodyタグが見つかりませんでした。");
                return;
            }

            // 処理開始
            ProcessNode(docBodyNode, bodyNode);

            // 抽出されたコンテンツを追加
            extractedContentList.Add(bodyNode.InnerHtml);
            Console.WriteLine($"抽出されたコンテンツのサイズ: {bodyNode.InnerHtml.Length} 文字");
        }

        /// <summary>
        /// ファイル全体を抽出する（抽出中のファイルで終了見出しがない場合）
        /// </summary>
        /// <param name="doc">HtmlDocument</param>
        private void ExtractEntireFile(HtmlDocument doc)
        {
            // 修正：テーブル構造とスタイルを保持するために、DOMツリーから直接抽出する方法に変更
            HtmlDocument extractedDoc = new HtmlDocument();
            HtmlNode bodyNode = extractedDoc.CreateElement("div");
            extractedDoc.DocumentNode.AppendChild(bodyNode);

            // BODYタグ以下の要素を対象に
            HtmlNode docBodyNode = doc.DocumentNode.SelectSingleNode("//body");
            if (docBodyNode == null)
            {
                Console.WriteLine("警告: bodyタグが見つかりませんでした。");
                return;
            }

            // 全てのノードを新しいドキュメントにコピー
            foreach (var node in docBodyNode.ChildNodes)
            {
                // ノードをコピー
                HtmlNode newNode = node.CloneNode(true);
                bodyNode.AppendChild(newNode);
            }

            // 抽出されたコンテンツを追加
            extractedContentList.Add(bodyNode.InnerHtml);
            Console.WriteLine($"抽出されたコンテンツのサイズ: {bodyNode.InnerHtml.Length} 文字");
        }

        /// <summary>
        /// ディレクトリ内のHTMLファイルを連番順にソートして取得
        /// </summary>
        /// <param name="directoryPath">ディレクトリのパス</param>
        /// <returns>ソートされたHTMLファイルパスのリスト</returns>
        private List<string> GetSortedHtmlFiles(string directoryPath)
        {
            string[] files = Directory.GetFiles(directoryPath, "*.html");
            
            // 連番でソート（01_html, 02_html, ...の形式を想定）
            var sortedFiles = files
                .Select(f => new 
                { 
                    FilePath = f,
                    SortKey = GetSortKey(Path.GetFileName(f)) 
                })
                .OrderBy(f => f.SortKey)
                .Select(f => f.FilePath)
                .ToList();

            return sortedFiles;
        }

        /// <summary>
        /// ファイル名からソートキーを取得（数値部分を抽出）
        /// </summary>
        /// <param name="fileName">ファイル名</param>
        /// <returns>ソートキー（数値）</returns>
        private int GetSortKey(string fileName)
        {
            // ファイル名から数字部分を抽出
            var match = Regex.Match(fileName, @"^(\d+)");
            if (match.Success && int.TryParse(match.Groups[1].Value, out int result))
            {
                return result;
            }
            return int.MaxValue; // 数字がない場合は最後にソート
        }

        /// <summary>
        /// 抽出したHTMLをファイルに保存する
        /// </summary>
        /// <param name="contentList">抽出したHTML内容のリスト</param>
        /// <param name="stylesList">抽出したスタイル情報のリスト</param>
        /// <param name="filePath">保存先ファイルパス</param>
        private void SaveToFile(List<string> contentList, List<string> stylesList, string filePath)
        {
            // ディレクトリが存在しない場合は作成
            string directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // HTML全体を構築
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html>");
            sb.AppendLine("<head>");
            sb.AppendLine("<meta charset=\"UTF-8\">");
            sb.AppendLine("<title>抽出された損益計算書</title>");
            
            // 抽出したスタイル情報をすべて追加
            foreach (var style in stylesList)
            {
                sb.AppendLine(style);
            }
            
            // デフォルトのスタイルも追加
            sb.AppendLine("<style>");
            sb.AppendLine("body { font-family: Arial, sans-serif; margin: 20px; }");
            sb.AppendLine("table { border-collapse: collapse; width: 100%; }");
            sb.AppendLine("th, td { border: 1px solid #ddd; padding: 8px; }");
            sb.AppendLine("tr:nth-child(even) { background-color: #f2f2f2; }");
            sb.AppendLine("</style>");
            
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");

            // 抽出したコンテンツを追加
            foreach (var content in contentList)
            {
                sb.AppendLine(content);
            }

            sb.AppendLine("</body>");
            sb.AppendLine("</html>");

            // ファイルに書き込み
            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        /// <summary>
        /// 単一のHTMLファイルから指定されたセクションを抽出する（元のメソッド - 互換性のために残す）
        /// </summary>
        /// <param name="htmlFilePath">HTMLファイルのパス</param>
        /// <param name="outputFilePath">出力ファイルのパス</param>
        /// <returns>抽出に成功したかどうか</returns>
        public bool ExtractAndSaveSection(string htmlFilePath, string outputFilePath)
        {
            try
            {
                // 状態をリセット
                extractionStarted = false;
                extractionCompleted = false;
                extractedContentList.Clear();
                extractedStylesList.Clear(); // スタイルリストもクリア

                // HTMLファイルを処理
                ProcessHtmlFile(htmlFilePath);

                // 抽出したHTMLを保存
                if (extractedContentList.Count > 0)
                {
                    SaveToFile(extractedContentList, extractedStylesList, outputFilePath);
                    Console.WriteLine($"セクションの抽出に成功しました。出力ファイル: {outputFilePath}");
                    return true;
                }
                else
                {
                    Console.WriteLine("指定されたセクションが見つかりませんでした。");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"エラーが発生しました: {ex.Message}");
                return false;
            }
        }
    }

    /// <summary>
    /// プログラムのエントリポイント
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("EDINETの有価証券報告書HTMLセクション抽出ツール");
            
            string path;
            string outputFilePath;
            bool isDirectory;

            if (args.Length >= 2)
            {
                // コマンドライン引数から入力パスと出力ファイルを取得
                path = args[0];
                outputFilePath = args[1];
                isDirectory = Directory.Exists(path);
            }
            else
            {
                // ユーザーに入力を求める
                Console.Write("単一ファイル抽出(F)かディレクトリ内の複数ファイル抽出(D)か選択してください [F/D]: ");
                string choice = Console.ReadLine().Trim().ToUpper();
                isDirectory = choice == "D";

                if (isDirectory)
                {
                    Console.Write("HTMLファイルが存在するディレクトリのパスを入力してください: ");
                    path = Console.ReadLine();
                }
                else
                {
                    Console.Write("入力HTMLファイルのパスを入力してください: ");
                    path = Console.ReadLine();
                }

                Console.Write("出力HTMLファイルのパスを入力してください: ");
                outputFilePath = Console.ReadLine();
            }

            // 入力パスの存在確認
            if (isDirectory && !Directory.Exists(path))
            {
                Console.WriteLine($"ディレクトリが見つかりません: {path}");
                Console.WriteLine("Enterキーを押して終了してください...");
                Console.ReadLine();
                return;
            }
            else if (!isDirectory && !File.Exists(path))
            {
                Console.WriteLine($"ファイルが見つかりません: {path}");
                Console.WriteLine("Enterキーを押して終了してください...");
                Console.ReadLine();
                return;
            }

            // セクション抽出実行
            HtmlSectionExtractor extractor = new HtmlSectionExtractor();
            bool success;

            if (isDirectory)
            {
                success = extractor.ExtractAndSaveSectionFromDirectory(path, outputFilePath);
            }
            else
            {
                success = extractor.ExtractAndSaveSection(path, outputFilePath);
            }

            if (success)
            {
                Console.WriteLine("処理が正常に完了しました。");
            }
            else
            {
                Console.WriteLine("処理に失敗しました。");
            }

            Console.WriteLine("Enterキーを押して終了してください...");
            Console.ReadLine();
        }
    }
}
