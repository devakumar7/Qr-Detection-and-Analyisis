QR Code Detection and Analysis

This Java program utilizes the Apache POI and BoofCV libraries to detect and analyze QR codes in a set of images. The detected information is stored in an Excel file for further analysis.
Features

    QR Code Detection: Detects QR codes in images using the BoofCV library.
    Bounding Box Drawing: Draws bounding boxes around detected QR codes on the original images.
    Image Analysis: Analyzes the detected QR codes for blur and pixel quality.
    Data Storage: Stores analysis results in an Excel file for easy reference.

Usage

    Input and Output Paths:
        Set the inputFolderPath to the directory containing images.
        Set the outputFolderPath to the directory where images with bounding boxes will be saved.
        Set the excelFilePath to the path where the Excel file will be created or updated.

    Allowed Image Extensions:
        Specify the allowed image file extensions in the ALLOWED_EXTENSIONS array.

    Run the Program:
        Execute the main method to process images and update the Excel file.

Dependencies

    Apache POI: Used for creating and updating Excel files.
    BoofCV: A computer vision library for Java.

Excel File Structure

    The program updates an Excel file with a sheet named "ImageRecords" containing the following columns:
        imageName: Name of the processed image.
        message: QR code message.
        laplacianVariance: Variance of Laplacian values for blur detection.
        brightness: Average brightness for pixel quality analysis.
        errorCorrectionLevel: QR code error correction level.
        height: Height of the detected QR code.
        width: Width of the detected QR code.
        heightXwidth: Combination of height and width.
        detectedAt*: Presence of the QR code at various resolutions (* represents resolution values).

Image Analysis

    Blur Detection:
        Uses Laplacian variance values to determine the sharpness of the QR code.

    Pixel Quality:
        Calculates the average brightness of pixels to assess the overall image quality.
