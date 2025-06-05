using System.IO;
using DocumentFormat.OpenXml.Packaging;
using Markdig;
using Markdig.Renderers.Docx;
using Microsoft.Extensions.Logging.Abstractions;

public static class MarkdownExtensions
{
    /// <summary>
    /// Converts a Markdown string into an in-memory DOCX document and returns a <see cref="MemoryStream"/> containing the result.
    /// If not specified, default document styles and a standard Markdown pipeline are used.
    /// The DOCX document is generated from a standard in-memory template.
    /// </summary>
    /// <param name="markdown">The Markdown text to convert.</param>
    /// <param name="styles">Optional styles to apply to the document (instance of <see cref="DocumentStyles"/>).</param>
    /// <param name="pipeline">Optional Markdig pipeline to use for Markdown parsing.</param>
    /// <returns>A <see cref="MemoryStream"/> containing the generated DOCX file.</returns>
    public static MemoryStream ToDocxStream(
    string markdown,
    DocumentStyles? styles = null,
    MarkdownPipeline? pipeline = null)
    {
        styles ??= new DocumentStyles();
        pipeline ??= new MarkdownPipelineBuilder().UseEmphasisExtras().Build();

        // Ottieni documento + stream associato
        var (document, stream) = DocxTemplateHelper.GetStandardTemplate();

        using(document)
        {
            var renderer = new DocxDocumentRenderer(document,styles,NullLogger<DocxDocumentRenderer>.Instance);
            Markdown.Convert(markdown,renderer,pipeline);
        }

        stream.Position = 0;
        return stream;
    }
}
