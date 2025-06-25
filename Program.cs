using HtmlAgilityPack;
using System.Text;

// Parse command line arguments
string? inputFile = null;
string? outputFile = null;

if (args.Length == 0)
{
    Console.WriteLine("Error: No arguments provided.");
    ShowHelp();
    return;
}

for (int i = 0; i < args.Length; i++)
{
    switch (args[i].ToLower())
    {
        case "-i":
        case "--input":
            if (i + 1 < args.Length)
            {
                inputFile = args[i + 1];
                i++; // Skip the next argument since we consumed it
            }
            else
            {
                Console.WriteLine("Error: --input requires a file path.");
                return;
            }
            break;
        case "-o":
        case "--output":
            if (i + 1 < args.Length)
            {
                outputFile = args[i + 1];
                i++; // Skip the next argument since we consumed it
            }
            else
            {
                Console.WriteLine("Error: --output requires a file path.");
                return;
            }
            break;
        case "-h":
        case "--help":
            ShowHelp();
            return;
        default:
            // If it doesn't start with -, treat as input file (for backward compatibility)
            if (!args[i].StartsWith("-"))
            {
                if (inputFile == null)
                {
                    inputFile = args[i];
                }
                else
                {
                    Console.WriteLine($"Error: Unexpected argument '{args[i]}'.");
                    ShowHelp();
                    return;
                }
            }
            else
            {
                Console.WriteLine($"Error: Unknown option '{args[i]}'.");
                ShowHelp();
                return;
            }
            break;
    }
}

// Validate required parameters
if (inputFile == null)
{
    Console.WriteLine("Error: Input file is required.");
    ShowHelp();
    return;
}

if (outputFile == null)
{
    Console.WriteLine("Error: Output file is required.");
    ShowHelp();
    return;
}

// Resolve full paths
var htmlPath = Path.IsPathRooted(inputFile) ? inputFile : Path.Combine(Environment.CurrentDirectory, inputFile);
var outputPath = Path.IsPathRooted(outputFile) ? outputFile : Path.Combine(Environment.CurrentDirectory, outputFile);

// Check if input file exists
if (!File.Exists(htmlPath))
{
    Console.WriteLine($"Error: HTML file not found: {htmlPath}");
    return;
}

Console.WriteLine($"Loading HTML file: {htmlPath}");
var doc = new HtmlDocument();
doc.Load(htmlPath);

// Find and extract content from the document content area, excluding UI elements
// Target scriptor-pageFrame elements which contain the actual document content
var contentElements = doc.DocumentNode.SelectNodes("//div[contains(@class, 'scriptor-pageFrame')]");

if (contentElements == null || contentElements.Count == 0)
{
    // Fallback to main element if scriptor-pageFrame not found
    var mainElement = doc.DocumentNode.SelectSingleNode("//main");
    if (mainElement == null)
    {
        throw new Exception("Content area not found in the HTML file.");
    }
    contentElements = new HtmlNodeCollection(mainElement.ParentNode) { mainElement };
}

var result = new StringBuilder();

// Process each content page frame
bool isFirstContent = true;
foreach (var contentElement in contentElements)
{
    ExtractTextFromNode(contentElement, result, 0, isFirstContent);
    isFirstContent = false;
}

var textContent = result.ToString();

// Write to output file
File.WriteAllText(outputPath, textContent, Encoding.UTF8);

Console.WriteLine($"Text extraction completed. Output saved to: {outputPath}");
Console.WriteLine($"Extracted {textContent.Split('\n', StringSplitOptions.RemoveEmptyEntries).Length} lines of text.");

static void ShowHelp()
{
    Console.WriteLine("LoopToMd - Extract meaningful content from Microsoft Loop HTML exports");
    Console.WriteLine();
    Console.WriteLine("Usage:");
    Console.WriteLine("  LoopToMd -i <input-file> -o <output-file>");
    Console.WriteLine("  LoopToMd --input <input-file> --output <output-file>");
    Console.WriteLine("  LoopToMd <input-file> -o <output-file>");
    Console.WriteLine();
    Console.WriteLine("Options:");
    Console.WriteLine("  -i, --input <file>     Input HTML file (required)");
    Console.WriteLine("  -o, --output <file>    Output markdown file (required)");
    Console.WriteLine("  -h, --help             Show this help message");
    Console.WriteLine();
    Console.WriteLine("Examples:");
    Console.WriteLine("  LoopToMd -i raw.html -o extracted_text.md");
    Console.WriteLine("  LoopToMd --input notes.html --output notes.md");
    Console.WriteLine("  LoopToMd myfile.html -o output.md");
}


void ExtractTextFromNode(HtmlNode node, StringBuilder result, int currentIndentLevel = 0, bool isFirstPage = false)
{
    // Skip UI elements and navigation components
    var classAttribute = node.GetAttributeValue("class", "");
    if (ShouldSkipNode(node, classAttribute))
    {
        return;
    }

    // Check if this might be the document title (first significant text on first page)
    if (isFirstPage && IsDocumentTitle(node, classAttribute))
    {
        var titleText = NormalizeWhitespace(node.InnerText);
        if (!string.IsNullOrWhiteSpace(titleText))
        {
            result.AppendLine($"# {titleText}");
            result.AppendLine(); // Add blank line after title
        }
        return;
    }

    // Check if this is a link element
    if (node.Name.ToLower() == "a" && !string.IsNullOrWhiteSpace(node.GetAttributeValue("href", "")))
    {
        var href = node.GetAttributeValue("href", "");
        var linkText = NormalizeWhitespace(node.InnerText);
        
        if (!string.IsNullOrWhiteSpace(linkText) && !string.IsNullOrWhiteSpace(href))
        {
            // Create markdown link format: [text](url)
            var markdownLink = $"[{linkText}]({href})";
            result.AppendLine(markdownLink);
        }
        return; // Don't process children since we got the full link
    }

    // Check if this is a heading element
    if (node.GetAttributeValue("role", "") == "heading")
    {
        var headingLevel = node.GetAttributeValue("aria-level", "1");
        var headingText = NormalizeWhitespace(node.InnerText);
        
        if (!string.IsNullOrWhiteSpace(headingText))
        {
            // Convert aria-level to markdown heading
            var markdownHeading = headingLevel switch
            {
                "1" => "## ",
                "2" => "### ",
                "3" => "#### ",
                _ => "" // Default to body for any other levels
            };
            
            result.AppendLine($"{markdownHeading}{headingText}");
            result.AppendLine(); // Add blank line after heading
        }
        return; // Don't process children since we got the full heading text
    }

    // Check if this is a list item
    if (node.Name.ToLower() == "li")
    {
        var listItemText = GetListItemText(node);
        if (!string.IsNullOrWhiteSpace(listItemText))
        {
            // Determine list type based on marker classes and indentation level
            var markerPrefix = GetListMarkerPrefix(node, currentIndentLevel);
            result.AppendLine($"{markerPrefix}{listItemText}");
        }
        return; // Don't process children since we got the full list item text
    }

    // Check if this is a list container (ul/ol) - we'll process the children but not add any text for the container itself
    if (node.Name.ToLower() == "ul" || node.Name.ToLower() == "ol")
    {
        // Just process children, don't add any text for the list container
        foreach (var child in node.ChildNodes)
        {
            if (child.NodeType == HtmlNodeType.Element)
            {
                ExtractTextFromNode(child, result, currentIndentLevel, isFirstPage);
            }
        }
        return;
    }

    // Check if this is a list item container div with margin-left styling
    if (node.GetAttributeValue("class", "").Contains("scriptor-listItem"))
    {
        var indentLevel = GetIndentationLevel(node);
        // Process children with the detected indentation level
        foreach (var child in node.ChildNodes)
        {
            if (child.NodeType == HtmlNodeType.Element)
            {
                ExtractTextFromNode(child, result, indentLevel, isFirstPage);
            }
        }
        return;
    }

    // If this is a leaf node with text content
    if (!node.HasChildNodes || node.ChildNodes.All(child => child.NodeType == HtmlNodeType.Text))
    {
        var text = NormalizeWhitespace(node.InnerText);
        if (!string.IsNullOrWhiteSpace(text))
        {
            result.AppendLine(text);
        }
    }
    else
    {
        // Recursively process child nodes
        foreach (var child in node.ChildNodes)
        {
            if (child.NodeType == HtmlNodeType.Element)
            {
                ExtractTextFromNode(child, result, currentIndentLevel, isFirstPage);
            }
        }
    }
}

string GetListItemText(HtmlNode listItemNode)
{
    var textParts = new List<string>();
    
    // Extract text from scriptor-textRun elements within the list item
    var textRuns = listItemNode.SelectNodes(".//span[contains(@class, 'scriptor-textRun')]");
    if (textRuns != null && textRuns.Count > 0)
    {
        foreach (var textRun in textRuns)
        {
            var text = NormalizeWhitespace(textRun.InnerText);
            if (!string.IsNullOrWhiteSpace(text))
            {
                textParts.Add(text);
            }
        }
    }
    
    // Also look for any links within the list item (they might be in sibling spans)
    var links = listItemNode.SelectNodes(".//a[@href]");
    if (links != null && links.Count > 0)
    {
        foreach (var link in links)
        {
            var href = link.GetAttributeValue("href", "");
            var linkText = NormalizeWhitespace(link.InnerText);
            
            if (!string.IsNullOrWhiteSpace(linkText) && !string.IsNullOrWhiteSpace(href))
            {
                textParts.Add($"[{linkText}]({href})");
            }
        }
    }
    
    if (textParts.Count > 0)
    {
        return string.Join(" ", textParts);
    }
    
    // Fallback to inner text if no textRuns or links found
    return NormalizeWhitespace(listItemNode.InnerText) ?? "";
}

string NormalizeWhitespace(string? text)
{
    if (string.IsNullOrWhiteSpace(text))
        return "";
    
    // Decode HTML entities (like &amp; -> &, &lt; -> <, etc.)
    var decodedText = System.Net.WebUtility.HtmlDecode(text);
    
    // Replace multiple whitespace characters (including newlines, tabs, spaces) with single spaces
    return System.Text.RegularExpressions.Regex.Replace(decodedText.Trim(), @"\s+", " ");
}

string GetListMarkerPrefix(HtmlNode listItemNode, int indentLevel)
{
    // Create indentation (2 spaces per level)
    var indent = new string(' ', indentLevel * 2);
    
    // Check for checkbox items by looking for checkbox marker elements within the list item
    var checkboxMarker = listItemNode.SelectSingleNode(".//span[contains(@class, 'scriptor-listItem-marker-checkbox')]");
    if (checkboxMarker != null)
    {
        var isChecked = GetCheckboxState(listItemNode);
        return isChecked ? $"{indent}- [x] " : $"{indent}- [ ] ";
    }
    
    // Check for bullet points by looking for bullet marker elements
    var bulletMarker = listItemNode.SelectSingleNode(".//span[contains(@class, 'scriptor-listItem-marker-bullet')] | .//*[contains(@class, 'scriptor-listItem-marker-bullet')]");
    if (bulletMarker != null)
    {
        return $"{indent}- "; // Markdown bullet point
    }
    
    // Check parent container for ordered vs unordered
    var parentList = listItemNode.ParentNode;
    if (parentList != null)
    {
        if (parentList.Name.ToLower() == "ol")
        {
            return $"{indent}1. "; // Markdown numbered list
        }
        else if (parentList.Name.ToLower() == "ul")
        {
            return $"{indent}- "; // Markdown bullet point
        }
    }
    
    // Default to bullet point
    return $"{indent}- ";
}

bool GetCheckboxState(HtmlNode listItemNode)
{
    // Look for checkbox marker element within this list item
    var checkboxMarker = listItemNode.SelectSingleNode(".//span[contains(@class, 'scriptor-listItem-marker-checkbox')]");
    if (checkboxMarker != null)
    {
        // Check aria-checked attribute
        var ariaChecked = checkboxMarker.GetAttributeValue("aria-checked", "false");
        if (ariaChecked == "true")
            return true;
        
        // Check class names for checked/unchecked state
        var markerClass = checkboxMarker.GetAttributeValue("class", "");
        if (markerClass.Contains("scriptor-listItem-marker-checkbox-checked"))
            return true;
        
        // Also check the text marker class
        var textMarker = checkboxMarker.SelectSingleNode(".//span[contains(@class, 'scriptor-listItem-marker-text-checkbox')]");
        if (textMarker != null)
        {
            var textMarkerClass = textMarker.GetAttributeValue("class", "");
            if (textMarkerClass.Contains("scriptor-listItem-marker-text-checkbox-checked"))
                return true;
        }
    }
    
    return false; // Default to unchecked
}

int GetIndentationLevel(HtmlNode listItemNode)
{
    var style = listItemNode.GetAttributeValue("style", "");
    
    // Extract margin-left value from style attribute
    var marginLeftMatch = System.Text.RegularExpressions.Regex.Match(style, @"margin-left:\s*(\d+)px");
    if (marginLeftMatch.Success && int.TryParse(marginLeftMatch.Groups[1].Value, out int marginLeft))
    {
        // Convert margin-left to indentation level
        // 27px = level 0, 54px = level 1, 81px = level 2, etc.
        return Math.Max(0, (marginLeft - 27) / 27);
    }
    
    return 0; // Default to no indentation
}

bool IsDocumentTitle(HtmlNode node, string classAttribute)
{
    // Check if this node contains significant text that could be the document title
    var innerText = NormalizeWhitespace(node.InnerText);
    
    // Skip empty nodes or very short text
    if (string.IsNullOrWhiteSpace(innerText) || innerText.Length < 3)
    {
        return false;
    }
    
    // Look for the actual document title within scriptor content
    // The document title appears as scriptor-textRun within scriptor-paragraph
    if (classAttribute.Contains("scriptor-textRun") && 
        classAttribute.Contains("scriptor-inline") &&
        node.ParentNode != null)
    {
        var parentClass = node.ParentNode.GetAttributeValue("class", "");
        if (parentClass.Contains("scriptor-paragraph"))
        {
            // This is likely the document title - it's the first significant scriptor content
            // Additional validation: ensure it's not part of a list or other structure
            var grandParent = node.ParentNode.ParentNode;
            if (grandParent != null)
            {
                var grandParentClass = grandParent.GetAttributeValue("class", "");
                // Make sure it's not part of a list structure
                if (!grandParentClass.Contains("scriptor-listItem"))
                {
                    return true;
                }
            }
        }
    }
    
    return false;
}

bool ShouldSkipNode(HtmlNode node, string classAttribute)
{
    // Skip Fluent UI components and generated CSS classes
    if (classAttribute.Contains("fui-") ||    // Fluent UI components
        classAttribute.Contains("___"))       // Generated CSS classes with triple underscore
    {
        return true;
    }
    
    // Skip SVG elements (icons)
    if (node.Name.ToLower() == "svg")
    {
        return true;
    }

    return false;
}

