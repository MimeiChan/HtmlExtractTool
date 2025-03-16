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
        // 抽出対象の終了見出し - 変更
        private const string END_SECTION_TITLE = "【株主資本等変動計算書】";

        // 抽出状態を管理する変数
        private bool extractionStarted = false;
        private bool extractionCompleted = false;
        private List<string> extractedContentList = new List<string>();
        
        // スタイル情報を保持する変数
        private HashSet<string> extractedStylesList = new HashSet<string>();

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
                extractedStylesList.Clear();

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
                    Console.WriteLine($"警告: {START_SECTION_TITLE}セクションは見つかりましたが、{END_SECTION_TITLE}セクションが見つかりませんでした。");
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
                    Console.WriteLine($"{START_SECTION_TITLE}セクションが見つかりませんでした。");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"エラーが発生しました: {ex.Message}");
                Console.WriteLine($"スタックトレース: {ex.StackTrace}");
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
                doc.OptionFixNestedTags = true;
                doc.OptionReadEncoding = true;
                doc.OptionCheckSyntax = false;
                doc.Load(htmlFilePath, Encoding.UTF8);

                // スタイル情報を抽出
                ExtractStyles(doc);

                // 可能性のある見出し要素を定義
                string[] headingTags = { "h1", "h2", "h3", "h4", "h5", "h6", "div", "span", "p" };

                // 開始ノードと終了ノードを検索
                HtmlNode startNode = null;
                HtmlNode endNode = null;

                // 見出しの検索
                foreach (var tag in headingTags)
                {
                    var nodes = doc.DocumentNode.SelectNodes($"//{tag}");
                    if (nodes == null) continue;

                    // 開始見出しの検索（まだ抽出が始まっていない場合）
                    if (!extractionStarted && startNode == null)
                    {
                        startNode = nodes.FirstOrDefault(n => n.InnerText.Contains(START_SECTION_TITLE));
                        if (startNode != null)
                        {
                            Console.WriteLine($"開始見出しが見つかりました: {startNode.InnerText}");
                        }
                    }

                    // 終了見出しの検索（常に検索する）
                    if (endNode == null)
                    {
                        endNode = nodes.FirstOrDefault(n => n.InnerText.Contains(END_SECTION_TITLE));
                        if (endNode != null)
                        {
                            Console.WriteLine($"終了見出しが見つかりました: {endNode.InnerText}");
                        }
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
                    ExtractEntireFile(doc);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ファイル処理中にエラーが発生しました: {htmlFilePath}, {ex.Message}");
                Console.WriteLine($"スタックトレース: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// HTMLドキュメントからスタイル情報を抽出する
        /// </summary>
        /// <param name="doc">HTMLドキュメント</param>
        private void ExtractStyles(HtmlDocument doc)
        {
            try
            {
                // headタグ内のstyleタグを抽出
                var styleNodes = doc.DocumentNode.SelectNodes("//head/style");
                if (styleNodes != null)
                {
                    foreach (var styleNode in styleNodes)
                    {
                        string styleContent = styleNode.InnerHtml;
                        if (!string.IsNullOrWhiteSpace(styleContent) && 
                            !extractedStylesList.Contains(styleContent))
                        {
                            extractedStylesList.Add(styleContent);
                            Console.WriteLine("スタイル情報を抽出しました。");
                        }
                    }
                }

                // headタグ内の最初のstyleタグ前にあるインラインスタイル定義を抽出
                var headNode = doc.DocumentNode.SelectSingleNode("//head");
                if (headNode != null)
                {
                    string headText = headNode.InnerText;
                    // CSSのようなスタイル定義を抽出する正規表現パターン
                    var cssPattern = @"\.[\w-]+ *\{ *[^}]+ *\}";
                    var cssMatches = Regex.Matches(headText, cssPattern);
                    
                    foreach (Match match in cssMatches)
                    {
                        string cssRule = match.Value;
                        if (!string.IsNullOrWhiteSpace(cssRule) && 
                            !extractedStylesList.Contains(cssRule))
                        {
                            extractedStylesList.Add(cssRule);
                            Console.WriteLine("インラインCSSルールを抽出しました。");
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
                        if (!string.IsNullOrWhiteSpace(href))
                        {
                            string linkOuterHtml = linkNode.OuterHtml;
                            if (!extractedStylesList.Contains(linkOuterHtml))
                            {
                                extractedStylesList.Add(linkOuterHtml);
                                Console.WriteLine("外部スタイルシートの参照を抽出しました。");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"スタイル抽出中にエラーが発生しました: {ex.Message}");
            }
        }

        /// <summary>
        /// XBRLマークアップを清掃する
        /// </summary>
        /// <param name="html">清掃対象のHTML</param>
        /// <returns>清掃後のHTML</returns>
        private string CleanXbrlMarkup(string html)
        {
            try
            {
                // <#text>タグを除去して内容を保持
                html = Regex.Replace(html, @"<\#text>(.*?)<\/\#text>", "$1", RegexOptions.Singleline);
                
                // ix:タグの処理
                html = Regex.Replace(html, @"<ix:[^>]*>(.*?)<\/ix:[^>]*>", "$1", RegexOptions.Singleline);
                
                return html;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"XBRL清掃中にエラーが発生しました: {ex.Message}");
                return html; // エラー時は元のHTMLを返す
            }
        }

        /// <summary>
        /// ノードを再帰的にコピーする (XBRL対応)
        /// </summary>
        private HtmlNode DeepCopyNode(HtmlNode originalNode, HtmlDocument targetDoc)
        {
            try
            {
                // 空白ノードやコメントノードは無視
                if (originalNode.NodeType == HtmlNodeType.Comment || 
                    (originalNode.NodeType == HtmlNodeType.Text && string.IsNullOrWhiteSpace(originalNode.InnerText)))
                {
                    return null;
                }

                // テキストノードの場合はXBRLマークアップを除去
                if (originalNode.NodeType == HtmlNodeType.Text)
                {
                    var textNode = targetDoc.CreateTextNode(CleanXbrlMarkup(originalNode.InnerText));
                    return textNode;
                }

                // 要素ノードの場合
                HtmlNode newNode = targetDoc.CreateElement(originalNode.Name);
                
                // 属性をコピー
                foreach (var attr in originalNode.Attributes)
                {
                    newNode.SetAttributeValue(attr.Name, attr.Value);
                }

                // 子ノードを再帰的にコピー
                if (originalNode.HasChildNodes)
                {
                    foreach (var child in originalNode.ChildNodes)
                    {
                        var copiedChild = DeepCopyNode(child, targetDoc);
                        if (copiedChild != null)
                        {
                            newNode.AppendChild(copiedChild);
                        }
                    }
                }
                else
                {
                    // 子ノードがない場合、InnerHtmlをクリーンアップして設定
                    if (!string.IsNullOrEmpty(originalNode.InnerHtml))
                    {
                        newNode.InnerHtml = CleanXbrlMarkup(originalNode.InnerHtml);
                    }
                }

                return newNode;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ノードコピー中にエラーが発生しました: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 開始ノードと終了ノードの間のコンテンツを抽出する（改善版）
        /// </summary>
        private void ExtractContentBetweenNodes(HtmlNode startNode, HtmlNode endNode, HtmlDocument doc)
        {
            try
            {
                // 抽出用のHTML文書を作成
                HtmlDocument extractedDoc = new HtmlDocument();
                HtmlNode rootNode = extractedDoc.CreateElement("div");
                extractedDoc.DocumentNode.AppendChild(rootNode);

                // 開始ノードと終了ノードの共通の親ノードを見つける
                HtmlNode commonAncestor = FindCommonAncestor(startNode, endNode, doc);
                if (commonAncestor == null)
                {
                    Console.WriteLine("共通の親ノードが見つかりませんでした。");
                    return;
                }

                Console.WriteLine($"共通の親ノード: {commonAncestor.Name}");

                bool isCapturing = false;
                ProcessNodesBetween(commonAncestor, startNode, endNode, extractedDoc, rootNode, ref isCapturing);

                // 結果を保存
                string extractedHtml = CleanXbrlMarkup(rootNode.InnerHtml);
                extractedContentList.Add(extractedHtml);
                Console.WriteLine($"抽出されたコンテンツのサイズ: {extractedHtml.Length} 文字");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"コンテンツ抽出中にエラーが発生しました: {ex.Message}");
                Console.WriteLine($"スタックトレース: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// 開始ノードと終了ノードの間のノードを処理する
        /// </summary>
        private void ProcessNodesBetween(HtmlNode currentNode, HtmlNode startNode, HtmlNode endNode, 
                                        HtmlDocument targetDoc, HtmlNode targetParent, ref bool isCapturing)
        {
            // 開始ノードに到達
            if (currentNode == startNode)
            {
                isCapturing = true;
                HtmlNode copiedNode = DeepCopyNode(currentNode, targetDoc);
                if (copiedNode != null)
                {
                    targetParent.AppendChild(copiedNode);
                }
                return; // 開始ノードの処理は完了
            }

            // 終了ノードに到達
            if (currentNode == endNode)
            {
                isCapturing = false;
                return; // 終了ノードは含めない
            }

            // 開始ノードと終了ノードの間のノードを捕捉
            if (isCapturing)
            {
                HtmlNode copiedNode = DeepCopyNode(currentNode, targetDoc);
                if (copiedNode != null)
                {
                    targetParent.AppendChild(copiedNode);
                }
                return;
            }

            // 子ノードに対して再帰
            if (currentNode.HasChildNodes)
            {
                foreach (var child in currentNode.ChildNodes)
                {
                    ProcessNodesBetween(child, startNode, endNode, targetDoc, targetParent, ref isCapturing);
                    if (!isCapturing && child == endNode)
                    {
                        break; // 終了ノードを超えたら処理終了
                    }
                }
            }
        }

        /// <summary>
        /// 2つのノードの共通の先祖を見つける
        /// </summary>
        private HtmlNode FindCommonAncestor(HtmlNode node1, HtmlNode node2, HtmlDocument doc)
        {
            // ノード1の先祖リストを作成
            var ancestors1 = new List<HtmlNode>();
            HtmlNode current = node1;
            while (current != null)
            {
                ancestors1.Add(current);
                current = current.ParentNode;
            }

            // ノード2の先祖をたどりながら、ノード1の先祖リストに含まれるか確認
            current = node2;
            while (current != null)
            {
                if (ancestors1.Contains(current))
                {
                    return current;
                }
                current = current.ParentNode;
            }

            // 共通の先祖が見つからない場合はbodyを返す
            return doc.DocumentNode.SelectSingleNode("//body");
        }

        /// <summary>
        /// 開始ノードから文書の最後までのコンテンツを抽出する（改善版）
        /// </summary>
        private void ExtractContentFromStartNode(HtmlNode startNode, HtmlDocument doc)
        {
            try
            {
                // 抽出用のHTML文書を作成
                HtmlDocument extractedDoc = new HtmlDocument();
                HtmlNode rootNode = extractedDoc.CreateElement("div");
                extractedDoc.DocumentNode.AppendChild(rootNode);

                // 開始ノードをコピー
                HtmlNode copiedStartNode = DeepCopyNode(startNode, extractedDoc);
                if (copiedStartNode != null)
                {
                    rootNode.AppendChild(copiedStartNode);
                }

                // 開始ノード以降のノードを抽出
                HtmlNode parent = startNode.ParentNode;
                bool foundStart = false;

                foreach (var sibling in parent.ChildNodes)
                {
                    if (sibling == startNode)
                    {
                        foundStart = true;
                        continue; // 開始ノードは既に追加済み
                    }

                    if (foundStart)
                    {
                        HtmlNode copiedNode = DeepCopyNode(sibling, extractedDoc);
                        if (copiedNode != null)
                        {
                            rootNode.AppendChild(copiedNode);
                        }
                    }
                }

                // 結果を保存
                string extractedHtml = CleanXbrlMarkup(rootNode.InnerHtml);
                extractedContentList.Add(extractedHtml);
                Console.WriteLine($"抽出されたコンテンツのサイズ: {extractedHtml.Length} 文字");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"開始ノードからのコンテンツ抽出中にエラーが発生しました: {ex.Message}");
                Console.WriteLine($"スタックトレース: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// 文書の先頭から終了ノードまでのコンテンツを抽出する（改善版）
        /// </summary>
        private void ExtractContentUntilEndNode(HtmlNode endNode, HtmlDocument doc)
        {
            try
            {
                // 抽出用のHTML文書を作成
                HtmlDocument extractedDoc = new HtmlDocument();
                HtmlNode rootNode = extractedDoc.CreateElement("div");
                extractedDoc.DocumentNode.AppendChild(rootNode);

                // 文書先頭からendNodeの直前までを抽出
                HtmlNode bodyNode = doc.DocumentNode.SelectSingleNode("//body");
                if (bodyNode == null)
                {
                    Console.WriteLine("警告: bodyタグが見つかりませんでした。");
                    return;
                }

                // 親ノードを取得
                HtmlNode parent = endNode.ParentNode;
                
                // 終了ノードの前のノードを全て抽出
                bool reachedEnd = false;
                foreach (var child in parent.ChildNodes)
                {
                    if (child == endNode)
                    {
                        reachedEnd = true;
                        break;
                    }

                    HtmlNode copiedNode = DeepCopyNode(child, extractedDoc);
                    if (copiedNode != null)
                    {
                        rootNode.AppendChild(copiedNode);
                    }
                }

                if (!reachedEnd)
                {
                    Console.WriteLine("警告: 終了ノードまでの抽出が完了しませんでした。");
                }

                // 結果を保存
                string extractedHtml = CleanXbrlMarkup(rootNode.InnerHtml);
                extractedContentList.Add(extractedHtml);
                Console.WriteLine($"抽出されたコンテンツのサイズ: {extractedHtml.Length} 文字");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"終了ノードまでのコンテンツ抽出中にエラーが発生しました: {ex.Message}");
                Console.WriteLine($"スタックトレース: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// ファイル全体を抽出する（抽出中のファイルで終了見出しがない場合）
        /// </summary>
        private void ExtractEntireFile(HtmlDocument doc)
        {
            try
            {
                // 抽出用のHTML文書を作成
                HtmlDocument extractedDoc = new HtmlDocument();
                HtmlNode rootNode = extractedDoc.CreateElement("div");
                extractedDoc.DocumentNode.AppendChild(rootNode);

                // BODYタグ以下の要素を対象に
                HtmlNode bodyNode = doc.DocumentNode.SelectSingleNode("//body");
                if (bodyNode == null)
                {
                    Console.WriteLine("警告: bodyタグが見つかりませんでした。");
                    return;
                }

                // 全てのノードをコピー
                foreach (var node in bodyNode.ChildNodes)
                {
                    HtmlNode copiedNode = DeepCopyNode(node, extractedDoc);
                    if (copiedNode != null)
                    {
                        rootNode.AppendChild(copiedNode);
                    }
                }

                // 結果を保存
                string extractedHtml = CleanXbrlMarkup(rootNode.InnerHtml);
                extractedContentList.Add(extractedHtml);
                Console.WriteLine($"抽出されたコンテンツのサイズ: {extractedHtml.Length} 文字");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ファイル全体の抽出中にエラーが発生しました: {ex.Message}");
                Console.WriteLine($"スタックトレース: {ex.StackTrace}");
            }
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
        private void SaveToFile(List<string> contentList, HashSet<string> stylesList, string filePath)
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
            sb.AppendLine("<style>");
            foreach (var style in stylesList)
            {
                sb.AppendLine(style);
            }
            
            // デフォルトのスタイルも追加
            sb.AppendLine("body { font-family: 'ＭＳ 明朝', serif; margin: 20px; }");
            sb.AppendLine("table { border-collapse: collapse; width: auto; margin-bottom: 20px; }");
            sb.AppendLine("th, td { border: 1px solid #ddd; padding: 8px; }");
            sb.AppendLine("tr:nth-child(even) { background-color: #f9f9f9; }");
            sb.AppendLine("tr[style*=\"background-color\"] { background-color: inherit !important; }");
            
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
                extractedStylesList.Clear();

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
            Console.WriteLine($"Version: 1.1.0 (styleKeepBranch)");
            
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
