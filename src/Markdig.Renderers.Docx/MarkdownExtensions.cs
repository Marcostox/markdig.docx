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
    string markdown, // Removed 'this Markdown _' to fix CS0721
    DocumentStyles? styles = null,
    MarkdownPipeline? pipeline = null)
    {
        // Usa il template standard (già su MemoryStream)
        var document = DocxTemplateHelper.Standard;
        styles ??= new DocumentStyles();

        var renderer = new DocxDocumentRenderer(document,styles,NullLogger<DocxDocumentRenderer>.Instance);
        pipeline ??= new MarkdownPipelineBuilder().UseEmphasisExtras().Build();

        // Renderizza il markdown
        Markdown.Convert(markdown,renderer,pipeline);

        // Ottieni lo stream sottostante; flush e chiudi il documento prima
        var stream = document.MainDocumentPart?.GetStream();
        if(stream == null)
        {
            throw new InvalidOperationException("MainDocumentPart stream is null.");
        }

        if(stream is not MemoryStream ms)
        {
            throw new InvalidCastException("The stream is not a MemoryStream.");
        }

        document.Close();

        ms.Position = 0; // resetta per la lettura
        return ms;
    }
}
