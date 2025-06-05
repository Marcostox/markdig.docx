using System.IO;
using DocumentFormat.OpenXml.Packaging;
using Markdig;
using Markdig.Renderers.Docx;
using Microsoft.Extensions.Logging.Abstractions;

public static class MarkdownExtensions
{
    /// <summary>
    /// Converte una stringa markdown in un DOCX in memoria, restituendo un MemoryStream.
    /// Il template standard e gli stili standard vengono usati se non specificati.
    /// </summary>
    /// <param name="markdown">Il testo Markdown da convertire</param>
    /// <param name="styles">Gli stili opzionali da applicare (DocumentStyles)</param>
    /// <param name="pipeline">La pipeline opzionale di Markdig</param>
    /// <returns>Un MemoryStream contenente il file DOCX generato</returns>
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
