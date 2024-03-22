# SwanSeaPrintingApp

## Overview

SwanSeaPrintingApp is a sophisticated tool designed for B2B drop shipping businesses to streamline their order fulfillment and shipping process. It expertly handles order details by reading them from PDF files, which can contain multiple orders forwarded by partnering businesses (e.g., Business X, Y, Z). The application uniquely caters to companies that fulfill orders and ship products directly to customers on behalf of other businesses. By automating the creation of printing labels for each individual order and generating invoices for the businesses placing the orders, the SwanSeaPrintingApp significantly enhances operational efficiency and accuracy.

## Key Features

- **PDF Order Processing**: Reads in order details from PDF files, each potentially containing multiple orders from different businesses.
- **Label Generation**: For each order, generates a distinct PDF with a printing label that includes essential shipping information such as the postage label, customs declaration label, total order cost, delivery address, and shop name.
- **Invoice Creation**: Automatically compiles and generates invoices for each business based on the total orders processed, streamlining the billing process.
- **User Interface**: Features a user-friendly interface developed with Tkinter, allowing for easy selection and processing of order files.

## How to Use

1. Launch the SwanSeaPrintingApp.
2. Through the application's user interface, select a list of 6 different files & folders. These files should contain the PDFs with order details.
3. Click on the "Generate" button to start the processing of order details.
4. The app will then create a printing label for each order and compile invoices for each business involved. The generated documents will be saved automatically to a predefined location or can be customized as needed.

## Requirements

To ensure the SwanSeaPrintingApp runs smoothly, the following technical requirements must be met:

- **Python Version**: Python 3.8 or higher.
- **Libraries**:
  - `Pandas` for data manipulation and analysis.
  - `ReportLab` for creating PDF documents.
  - `Pillow` (PILLOW) for image processing within PDFs.
  - `PyPDF2` for reading and manipulating PDF files.
