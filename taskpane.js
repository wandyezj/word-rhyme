async function run() {
    return Word.run(async (context) => {
        // insert a paragraph at the end of the document.
        const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
        // change the paragraph color to blue.
        paragraph.font.color = "blue";
        await context.sync();
    });
}
