# pdf_to_excel

**pdf_to_excel** - this is a convenient way to transfer data from a Pdf file to an Excel table in a couple of minutes, completely free of charge.

**PdfToExcel** uses the best tools on the market to convert Pdf files. You can read about it [here](https://developer.adobe.com/document-services/docs/overview/).

**API documentation** - [https://developer.adobe.com/document-services](https://developer.adobe.com/document-services/docs/overview/pdf-services-api/)

## How to start?

* Your first step is to download this repository and build it using [Visual Studio](https://visualstudio.microsoft.com/).

* After installation, you need to go to the official [Adobe developers](https://developer.adobe.com/) website and create an account.

* Next you need to go to the console tab -> projects -> create a new project.

* Add an API to your project and configure it (pay attention to the photo).
	
    <div id="" align="center">
	<img src="https://github.com/OneCbyte/pdf_to_excel/blob/master/1.jpg?raw=true" width="550"/>
	<img src="https://github.com/OneCbyte/pdf_to_excel/blob/master/2.jpg?raw=true" width="540"/>
	
	<div align="center">Downloading the archive with the key will begin automatically after creating the project</div>
	</div>

* Rebuild the application by pressing Ctrl+Alt+F7 and go to the path ```\bin\Debug\net6.0-windows``` and transfer the private.key installed in advance there.

* In the same directory create a ```file.json``` file like this:
```json
{
  "client_credentials": {
    "client_id": "your client_id from Adobe project",
    "client_secret": "your client_secret from Adobe project"
  },
  "service_account_credentials": {
    "organization_id": "your organization_id from Adobe project",
    "account_id": "your account_id from Adobe project",
    "private_key_file": "\\private.key"
  }
}
```
Supply your data from Adobe project

* Finally, save the changes and run the application)

## Usage

To use pdf_to_excel you need to enter the path to the pdf file and come up with a name to save the output in excel. If you do not specify an output file name, it will be saved as ```output.xlsx```. 

The output is saved to the desktop.

## Warning!

Adobe API allows you to use your product for free about 500 times per project. After this period expires you will have to fetch new data and replace the ```file.json``` & ```private.key``` files.

It is also necessary to remember that to create an output file, the name must be unique and contain the ```.xlsx``` suffix, otherwise an error will be displayed.

It is also necessary to specify the full path to the pdf file: ```C:\project\pdf\1.pdf```, no need to use any separating characters.

In any case, in case of errors you will be notified.

___

<h2 align="center">Thanks for using my software!❤️</h2>
