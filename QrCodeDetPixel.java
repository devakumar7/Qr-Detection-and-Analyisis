package com.ctel.excel;

import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.imageio.ImageIO;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import boofcv.abst.fiducial.QrCodeDetector;
import boofcv.alg.fiducial.qrcode.QrCode;
import boofcv.factory.fiducial.FactoryFiducial;
import boofcv.io.image.ConvertBufferedImage;
import boofcv.struct.image.GrayU8;
import georegression.struct.point.Point2D_F64;

public class QrCodeDetPixel {

	private static String inputFolderPath = "C:\\Deva\\PythonWorkspace\\sirsImages\\checking\\temp";
	private static String outputFolderPath = "C:\\Deva\\PythonWorkspace\\sirsImages\\checking\\detected";
	private static String excelFilePath = "C:\\Deva\\PythonWorkspace\\sirsImages\\checking\\temp.xls";

	// Specify the list of allowed file extensions
	private static final String[] ALLOWED_EXTENSIONS = { ".jpg", ".png", ".bmp", ".gif" };

	public static void main(String[] args) throws IOException {
		// Create or open the Excel file outside of the loop
		Workbook workbook=getOrCreateWorkbook(excelFilePath);

		File inputFolder = new File(inputFolderPath);
		File[] imageFiles = inputFolder.listFiles();

		if (imageFiles != null) {
			for (File imageFile : imageFiles) {
				if (imageFile.isFile() && isAllowedExtension(imageFile.getName())) {
					processImage(imageFile, workbook); // Pass the workbook to the processImage method
				}
			}
		} else {
			System.err.println("Failed to list files in the input folder.");
		}

		// Close the workbook and save it after processing all images
		closeAndSaveWorkbook(workbook);


	}

	public static Workbook getOrCreateWorkbook(String filePath) throws IOException {
		Workbook workbook;

		// Check if the Excel file already exists at the specified path
		File excelFile = new File(filePath);

		if (excelFile.exists()) {
			// If the file exists, read the existing workbook
			FileInputStream fis = new FileInputStream(excelFile);
			workbook = new HSSFWorkbook(fis);
			fis.close();
		} else {
			// If the file doesn't exist, create a new workbook
			workbook = new HSSFWorkbook();
			FileOutputStream fos = new FileOutputStream(filePath);
			workbook.write(fos);
			fos.close();
		}

		return workbook;
	}

	private static boolean isAllowedExtension(String fileName) {
		for (String extension : ALLOWED_EXTENSIONS) {
			if (fileName.toLowerCase().endsWith(extension)) {
				return true;
			}
		}
		return false;
	}

	public static void writeEntityToExcel(Entity entity, String filePath, Workbook workbook) {
		try {
			// Create a sheet (or get the existing one) named after the imageName
			// Try to get the sheet with the given imageName
	        Sheet sheet = workbook.getSheet("ImageRecords");

			if (sheet == null) {
				sheet = workbook.createSheet("ImageRecords");
				// Create the header row
				Row headerRow = sheet.createRow(0);
				headerRow.createCell(2).setCellValue("laplacianVariance");
				headerRow.createCell(3).setCellValue("brightness");
				headerRow.createCell(1).setCellValue("message");
				headerRow.createCell(0).setCellValue("imageName");
				headerRow.createCell(4).setCellValue("errorCorrectionLevel");
				headerRow.createCell(5).setCellValue("height"); // Add this line
				headerRow.createCell(6).setCellValue("width"); // Add this line
				headerRow.createCell(7).setCellValue("heightXwidth"); // Add this line
				headerRow.createCell(8).setCellValue("detectedAt110");
				headerRow.createCell(9).setCellValue("detectedAt120");
				headerRow.createCell(10).setCellValue("detectedAt125");
				headerRow.createCell(11).setCellValue("detectedAt130");
				headerRow.createCell(12).setCellValue("detectedAt140");
				headerRow.createCell(13).setCellValue("detectedAt150");
				headerRow.createCell(14).setCellValue("detectedAt160");
				headerRow.createCell(15).setCellValue("detectedAt170");
				headerRow.createCell(16).setCellValue("detectedAt175");
				headerRow.createCell(17).setCellValue("detectedAt180");
				headerRow.createCell(18).setCellValue("detectedAt190");
				headerRow.createCell(19).setCellValue("detectedAt200");
				headerRow.createCell(20).setCellValue("detectedAt210");
				headerRow.createCell(21).setCellValue("detectedAt220");
				headerRow.createCell(22).setCellValue("detectedAt225");
				headerRow.createCell(23).setCellValue("detectedAt230");
				headerRow.createCell(24).setCellValue("detectedAt240");
				headerRow.createCell(25).setCellValue("detectedAt250");

			}

			// Create a new row for the entity data
			int rowCount = sheet.getLastRowNum() + 1;
			Row dataRow = sheet.createRow(rowCount);
			dataRow.createCell(2).setCellValue(entity.getLaplacianVariance());
			dataRow.createCell(3).setCellValue(entity.getBrightness());
			dataRow.createCell(1).setCellValue(entity.getMessage());
			dataRow.createCell(0).setCellValue(entity.getImageName());
			dataRow.createCell(4).setCellValue(entity.getErrorCorrectionLevel());
			dataRow.createCell(5).setCellValue(entity.getHeight());
			dataRow.createCell(6).setCellValue(entity.getWidth());
			dataRow.createCell(7).setCellValue(entity.getHeightXwidth());

			// Autosize columns for better readability
			for (int i = 0; i < 8; i++) {
				sheet.autoSizeColumn(i);
			}

			// Write the workbook back to the file
			FileOutputStream fos = new FileOutputStream(filePath);
			workbook.write(fos);
			fos.close();

			System.out.println("Data written to Excel successfully.");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static void processImage(File imageFile, Workbook workbook) {
		String filenameWithoutExtension = getFilenameWithoutExtension(imageFile.getName());

		try {

			BufferedImage qrCodeImage = ImageIO.read(imageFile);

			// Convert the BufferedImage to GrayU8 image
			GrayU8 grayImage = ConvertBufferedImage.convertFrom(qrCodeImage, (GrayU8) null);

			// Create an instance of QrCodeDetector
			QrCodeDetector<GrayU8> detector = FactoryFiducial.qrcode(null, GrayU8.class);

			// Detect and decode the QR codes in the image
			detector.process(grayImage);

			// Retrieve the list of detected QR codes
			List<QrCode> qrCodes = detector.getDetections();
			System.out.println(qrCodes.size());

			// Draw bounding boxes on the original image and save it
			drawBoundingBoxes(qrCodeImage, qrCodes, filenameWithoutExtension, workbook);

//			ImageIO.write(qrCodeImage, "png", new File(outputFolderPath, filenameWithoutExtension+"bounded"+".png"));

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	// Utility function to get the filename without extension
	private static String getFilenameWithoutExtension(String fileName) {
		int lastDotIndex = fileName.lastIndexOf(".");
		if (lastDotIndex > 0) {
			return fileName.substring(0, lastDotIndex);
		} else {
			return fileName;
		}
	}

	private static void drawBoundingBoxes(BufferedImage image, List<QrCode> qrCodes, String filenameWithoutExtension,
			Workbook workbook) {
		int counter = 1; // Initialize a counter for generating unique filenames
		Graphics2D g2 = image.createGraphics();
		g2.setStroke(new BasicStroke(3));
		g2.setColor(Color.GREEN);
		Entity dummyEntity = new Entity();
		dummyEntity.setImageName(filenameWithoutExtension);

		for (QrCode qrCode : qrCodes) {
			// Get the vertex points of the QR code polygon
			List<Point2D_F64> points = qrCode.bounds.vertexes.toList();
			if (points.size() == 4) {
				// Calculate the bounding box using the corner points of the polygon
				int minX = Integer.MAX_VALUE;
				int minY = Integer.MAX_VALUE;
				int maxX = Integer.MIN_VALUE;
				int maxY = Integer.MIN_VALUE;

				for (Point2D_F64 point : points) {
					int x = (int) point.x;
					int y = (int) point.y;
					minX = Math.min(minX, x);
					minY = Math.min(minY, y);
					maxX = Math.max(maxX, x);
					maxY = Math.max(maxY, y);
				}
				
				// Calculate height and width
		        int qrCodeWidth = maxX - minX;
		        int qrCodeHeight = maxY - minY;
		        
		        
		        
		        // Print height and width
		        System.out.println("Width: " + qrCodeWidth);
		        System.out.println(String.valueOf(qrCodeWidth));
		        System.out.println("Height: " + qrCodeHeight);
		        System.out.println(String.valueOf(qrCodeHeight));

//				 //Draw the bounding box
//				g2.setColor(Color.GREEN);
				Rectangle2D.Double boundingBox = new Rectangle2D.Double(minX, minY, maxX - minX, maxY - minY);

				// Generate a unique filename for each cropped image
				String filename = filenameWithoutExtension + "_cropped_" + counter + ".png";

				// Create a cropped image from the bounding box region
				BufferedImage croppedImage = image.getSubimage(minX, minY, maxX - minX, maxY - minY);

				// Check for blur and pixel quality for the entire image
				double lapVar = isBlurry(croppedImage);
				double brightness = hasGoodPixelQuality(croppedImage);

				System.out.println("Laplacian Variance Value : " + lapVar);
				System.out.println("Brightness Value : " + brightness);

				System.out.println("Message from : " + filename + " is : " + qrCode.message);
//				System.out.println("Error Correction Level for " + filename + " is : " + qrCode.error);
				System.out.println("Error Correction Level for " + filename + " is : " + qrCode.error.toString());

				// Create a dummy Entity object with sample data
				dummyEntity.setLaplacianVariance(lapVar);
				dummyEntity.setBrightness(brightness);
				dummyEntity.setMessage(qrCode.message);
				dummyEntity.setErrorCorrectionLevel(qrCode.error.toString());
//				dummyEntity.setHeight(qrCode.hei);
//				dummyEntity.setWidth(qrCodeWidth);
//				dummyEntity.setHeightXwidth(filenameWithoutExtension);

				System.out.println(dummyEntity.getImageName());

				// Save the cropped image
				try {
					ImageIO.write(croppedImage, "png", new File(outputFolderPath, filename));
				} catch (IOException e) {
					e.printStackTrace();
				}

//				g2.draw(boundingBox);
				// System.out.println("\nMessage: " + qrCode.message);

			}
			counter++;
		}
		g2.dispose();

		// Call the writeEntityToExcel method with the dummy data and file path
		writeEntityToExcel(dummyEntity, excelFilePath, workbook);

	}

	private static void closeAndSaveWorkbook(Workbook workbook) {
		try {
			FileOutputStream fos = new FileOutputStream(excelFilePath); // Specify the output Excel file
			workbook.write(fos);
			fos.close();
			System.out.println("Data written to Excel successfully.");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static double isBlurry(BufferedImage image) {

		GrayU8 qrCodeImgProcessing = ConvertBufferedImage.convertFrom(image, (GrayU8) null);

		int width = qrCodeImgProcessing.getWidth();
		int height = qrCodeImgProcessing.getHeight();

		// Create a matrix to store the Laplacian result as 64-bit floating point
		double[][] laplacianMatrix = new double[width][height];

		// Apply the Laplacian filter
		for (int y = 1; y < height - 1; y++) {
			for (int x = 1; x < width - 1; x++) {
				int pixel = image.getRGB(x, y) & 0xFF; // Get the grayscale value

				// Apply the Laplacian filter
				double laplacianValue = 4 * pixel;
				laplacianValue -= image.getRGB(x - 1, y) & 0xFF;
				laplacianValue -= image.getRGB(x + 1, y) & 0xFF;
				laplacianValue -= image.getRGB(x, y - 1) & 0xFF;
				laplacianValue -= image.getRGB(x, y + 1) & 0xFF;

				laplacianMatrix[x][y] = laplacianValue;
			}
		}

		// Calculate the mean of the Laplacian values
		double mean = calculateMean(laplacianMatrix);

		// Calculate the variance
		double variance = calculateVariance(laplacianMatrix, mean);

		System.out.println("Variance of Laplacian values: " + variance);

		// Adjust this threshold as needed
		int blurThreshold = 2000;

//        return variance < blurThreshold;

		return variance;

	}

	// Function to calculate the mean of Laplacian values
	private static double calculateMean(double[][] matrix) {
		int width = matrix.length;
		int height = matrix[0].length;
		double sum = 0.0;

		for (int y = 0; y < height; y++) {
			for (int x = 0; x < width; x++) {
				sum += matrix[x][y];
			}
		}

		return sum / (width * height);
	}

	// Function to calculate the variance of Laplacian values
	private static double calculateVariance(double[][] matrix, double mean) {
		int width = matrix.length;
		int height = matrix[0].length;
		double sumOfSquaredDifferences = 0.0;

		for (int y = 0; y < height; y++) {
			for (int x = 0; x < width; x++) {
				double difference = matrix[x][y] - mean;
				sumOfSquaredDifferences += difference * difference;
			}
		}

		return sumOfSquaredDifferences / (width * height);
	}

	public static double hasGoodPixelQuality(BufferedImage image) {

		double minBrightness = 50;
		double maxDarkness = 150;

		int width = image.getWidth();
		int height = image.getHeight();

		double totalBrightness = 0;

		for (int y = 0; y < height; y++) {
			for (int x = 0; x < width; x++) {
				int rgb = image.getRGB(x, y);
				int red = (rgb >> 16) & 0xFF;
				int green = (rgb >> 8) & 0xFF;
				int blue = rgb & 0xFF;

				// Calculate pixel brightness (you can use different weightings for RGB
				// channels)
				double pixelBrightness = (red + green + blue) / 3.0;
				totalBrightness += pixelBrightness;
			}
		}

		// Calculate the average brightness for all pixels
		double averageBrightness = totalBrightness / (width * height);

		System.out.println("Calculated Mean Brightness Value is : " + averageBrightness);

//		return averageBrightness >= minBrightness && averageBrightness <= maxDarkness;
		return averageBrightness;
	}

}