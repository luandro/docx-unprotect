# DocxUnprotect
Tool to unlock protected docx files that can't otherwise be edited.

## Global Installation

To install DocxUnprotect globally on your system, ensuring that you have Node.js installed, run:

```bash
npm install -g docx-unprotect
```

## Usage

To unprotect a DOCX file using the globally installed package, run the following command with the path to the protected DOCX file. Optionally, you can also provide the path where the unprotected DOCX file should be saved:

```bash
docx-unprotect <path-to-protected-docx> [path-to-unprotected-docx]
```

If you do not provide the path for the unprotected DOCX file, the script will automatically create an unprotected version in the same directory as the protected file.

## Debugging

If you encounter issues, you can enable debugging output by setting the `DEBUG` environment variable:

```bash
DEBUG=true node unprotectDocx.js
```

This will print additional information to the console during the execution of the script.

## License

This project is open-source and available under the MIT License.
