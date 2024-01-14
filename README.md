# Word Document Builder

This is a simple C# application that creates a Word document using a template and adds some text to it.

## How it works

The application uses the `Microsoft.Office.Interop.Word` namespace to interact with Microsoft Word. It creates a new Word application, opens a template document, and adds some text to it.

Here's a brief overview of what the code does:

1. It creates a new Word application.
2. It opens a template document specified by the `oTemplate` object.
3. It finds a bookmark named "oBookMark1" in the document and replaces it with the text "Some Text Here".
4. It inserts a new paragraph at the end of the document with the text "Heading 2".

## How to run the program

To run the program, you need to have Microsoft Word installed on your machine. You also need to replace the path in the `oTemplate` object with the path to your template document.

Once you've done that, you can run the program by clicking the "Run" button in your IDE or by using the `dotnet run` command in the terminal.

## Troubleshooting

If you encounter any issues while running the program, make sure that the path to your template document is correct and that the document contains a bookmark named "oBookMark1". If the issue persists, please open a new issue in this repository.
