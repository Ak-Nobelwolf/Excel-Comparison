package compare;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Properties;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import javax.servlet.annotation.MultipartConfig;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Servlet implementation class Compare
 */
@MultipartConfig
public class Compare extends HttpServlet {
	private static final long serialVersionUID = 1L;
	private static final Logger logger = LogManager.getLogger(Compare.class);

	public void init() {
		System.setProperty("log4j2.debug", "true");
		System.setProperty("java.io.tmpdir", "CompareUI/src/main/webapp/temp");
	}

	/**
	 * Default constructor.
	 */
	public Compare() {
		super();
	}

	/**
	 * Handles the POST request for comparing Excel files.
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		String methodName = "doPost";
		logger.debug("Entering {} method", methodName);

		// Get uploaded configuration file
		InputStream configStream = request.getPart("ConfigFile").getInputStream();

		// Load configuration file contents into Properties object
		Properties config = new Properties();

		Workbook workbook1 = null;
		Workbook workbook2 = null;

		try {
			InputStream sourceStream = request.getPart("SourceFile").getInputStream();
			InputStream targetStream = request.getPart("TargetFile").getInputStream();

			workbook1 = createWorkbook(sourceStream, request.getPart("SourceFile").getSubmittedFileName());
			workbook2 = createWorkbook(targetStream, request.getPart("TargetFile").getSubmittedFileName());

			// Compare the sheets and write the result to an output file
			File outputFile = new File("result_file.txt");

			try (PrintWriter writer = new PrintWriter(outputFile)) {
				// Calling Compare Sheets Method
				CompareSheets(workbook1, workbook2, writer, config, configStream);
			}

			response.setHeader("Content-Type", "text/plain");
			response.setHeader("Cache-Control", "no-cache,private,must-revalidate");
			response.setHeader("Pragma", "no-cache");
			response.setDateHeader("Expires", 0);
			response.setContentType("application/octet-stream");
			response.setHeader("Content-Disposition", "attachment;filename=\"result_file.txt\"");

			FileInputStream fileInputStream = new FileInputStream(outputFile);
			int i;

			while ((i = fileInputStream.read()) != -1) {
				response.getWriter().write(i);
			}
			fileInputStream.close();

		} catch (IOException ioe) {
			response.sendError(HttpServletResponse.SC_INTERNAL_SERVER_ERROR,
					"Internal Server error: " + ioe.getMessage());
			logger.error("IOException occurred: {}", ioe.getMessage());
		} catch (ServletException se) {
			response.sendError(HttpServletResponse.SC_INTERNAL_SERVER_ERROR,
					"Internal Server error: " + se.getMessage());
			logger.error("ServletException occurred: {}", se.getMessage());
		} catch (Exception e) {
			response.sendError(HttpServletResponse.SC_INTERNAL_SERVER_ERROR,
					"Internal Server error: " + e.getMessage());
			logger.error("Exception occurred: {}", e.getMessage());
		} finally {
			// Close the workbooks if they were opened
			try {
				if (workbook1 != null) {
					workbook1.close();
				}
				if (workbook2 != null) {
					workbook2.close();
				}
			} catch (IOException ioe) {
				// log close error
				logger.error("Error Closing workbooks: {}", ioe.getMessage());
			}
		}
		logger.debug("Exiting {} method", methodName);
		logger.info("Code Execution Completed");

	}

	private Workbook createWorkbook(InputStream inputStream, String fileName) throws IOException {
		Workbook workbook = null;
		if (fileName.endsWith(".xlsx")) {
			workbook = new XSSFWorkbook(inputStream);
		} else if (fileName.endsWith(".xls")) {
			workbook = new HSSFWorkbook(inputStream);
		} else {
			throw new IllegalArgumentException("Invalid File Format: " + fileName);
		}
		return workbook;
	}

	/**
	 * Compares the sheets of two Excel workbooks.
	 */
	private void CompareSheets(Workbook workbook1, Workbook workbook2, PrintWriter writer, Properties config,
			InputStream configStream) throws FileNotFoundException, IOException, NullPointerException {
		String methodName = "CompareSheets";
		logger.debug("Entering {} method", methodName);
		try {
			// Compare the sheets in the two workbooks
			for (int s = 0; s < workbook1.getNumberOfSheets(); s++) {

				Sheet sheet1 = workbook1.getSheetAt(s);
				Sheet sheet2 = workbook2.getSheetAt(s);

				// Create a map to store the column headings
				HashMap<String, Integer> map1 = new HashMap<>();
				HashMap<String, Integer> map2 = new HashMap<>();

				// Populate the map for the first workbook
				Row row1 = sheet1.getRow(0);
				if (row1 == null || row1.getPhysicalNumberOfCells() == 0) {
					// Handle null or empty row
					logger.warn("The first row in sheet1 is null or empty.");
					return;
				}
				int cellCount1 = row1.getPhysicalNumberOfCells();
				for (int i = 0; i < cellCount1; i++) {
					Cell cell = row1.getCell(i);
					String value = cell.getStringCellValue();
					map1.put(value, i);
				}

				// Populate the map for the second workbook
				Row row2 = sheet2.getRow(0);
				if (row2 == null || row2.getPhysicalNumberOfCells() == 0) {
					// Handle null or empty row
					logger.warn("The first row in sheet2 is null or empty.");
					return;
				}
				int cellCount2 = row2.getPhysicalNumberOfCells();
				for (int i = 0; i < cellCount2; i++) {
					Cell cell = row2.getCell(i);
					String value = cell.getStringCellValue();
					map2.put(value, i);
				}

				// Load the properties file
				config.load(configStream);

				// Read the column headings from the properties file
				String[] sourceColumns = config.getProperty("sourceColumns").split(",");
				String[] targetColumns = config.getProperty("targetColumns").split(",");

				// Check if the required columns exist in both workbooks
				for (String column : sourceColumns) {
					if (!map1.containsKey(column)) {
						writer.println("The Column '" + column + "' is not present in the Source file.");
						logger.warn("The Column '{}' is not present in the Source file.", column);
						return;
					}
				}

				for (String column : targetColumns) {
					if (!map2.containsKey(column)) {
						writer.println("The Column '" + column + "' is not present in the Target file.");
						logger.warn("The Column '{}' is not present in the Target file.", column);
						return;
					}
				}

				// Compare the values in each row
				Iterator<Row> iterator1 = sheet1.iterator();
				Iterator<Row> iterator2 = sheet2.iterator();

				// Skip the first row (as they are column headings)
				iterator1.next();
				iterator2.next();

				int rowNum = 1;
				while (iterator1.hasNext() && iterator2.hasNext()) {
					Row row3 = iterator1.next();
					Row row4 = iterator2.next();
					rowNum++;

					for (int j = 0; j < sourceColumns.length; j++) {
						Cell cell1 = row3.getCell(map1.get(sourceColumns[j]));
						Cell cell2 = row4.getCell(map2.get(targetColumns[j]));

						// Check if both cells are empty, if yes, skip comparison
						if (cell1 == null && cell2 == null) {
							continue;
						}

						// Check if one of the cells is empty, if yes, handle the case
						if (cell1 == null || cell2 == null) {
							writer.println("One of the cells is null in row " + rowNum + ", cell " + sourceColumns[j]
									+ " and " + targetColumns[j]);
							logger.warn("One of the cells is null in row {}, cell {} and {}", rowNum, sourceColumns[j],
									targetColumns[j]);
							continue;
						}

						// Check if both of the cells are empty, if yes, continue to next iteration
						if (cell1.getCellType() == CellType.BLANK && cell2.getCellType() == CellType.BLANK) {
							continue;
						}

						// Proceed with the comparison
						if (cell1.getCellType() == CellType.STRING && cell2.getCellType() == CellType.STRING) {
							// Compare String Values
							if (!cell1.getStringCellValue().equals(cell2.getStringCellValue())) {
								writer.println("Mismatch found in row " + rowNum + " at cell "
										+ cell1.getAddress().formatAsString() + ": " + "Source value is "
										+ cell1.getStringCellValue() + " and the Target Value is "
										+ cell2.getStringCellValue());
								logger.info(
										"Mismatch found in row {} at cell {}: Source value is '{}' and the Target Value is '{}'",
										rowNum, cell1.getAddress().formatAsString(), cell1.getStringCellValue(),
										cell2.getStringCellValue());
							}
						} else if (cell1.getCellType() == CellType.NUMERIC && cell2.getCellType() == CellType.NUMERIC) {
							// Compare numeric values
							if (cell1.getNumericCellValue() != cell2.getNumericCellValue()) {
								writer.println("Mismatch found in row " + rowNum + " at cell "
										+ cell1.getAddress().formatAsString() + ": " + "Source value is "
										+ cell1.getNumericCellValue() + " and the Target Value is "
										+ cell2.getNumericCellValue());
								logger.info(
										"Mismatch found in row {} at cell {}: Source value is {} and the Target Value is {}",
										rowNum, cell1.getAddress().formatAsString(), cell1.getNumericCellValue(),
										cell2.getNumericCellValue());
							}
						} else if (cell1.getCellType() == CellType.BOOLEAN && cell2.getCellType() == CellType.BOOLEAN) {
							// Compare boolean values
							if (cell1.getBooleanCellValue() != cell2.getBooleanCellValue()) {
								writer.println("Mismatch found in row " + rowNum + " at cell "
										+ cell1.getAddress().formatAsString() + ": " + "Source value is "
										+ cell1.getBooleanCellValue() + " and the Target Value is "
										+ cell2.getBooleanCellValue());
								logger.info(
										"Mismatch found in row {} at cell {}: Source value is {} and the Target Value is {}",
										rowNum, cell1.getAddress().formatAsString(), cell1.getBooleanCellValue(),
										cell2.getBooleanCellValue());
							}
						}
					}
				}
			}
		} catch (FileNotFoundException fnfe) {
			logger.error("Found a FileNotFoundException: {}", fnfe);
		} catch (IOException ioe) {
			logger.error("Found an IOException: {}", ioe);
		} catch (NullPointerException npe) {
			logger.error("Found a NullPointerException: {}", npe);
		} catch (Exception e) {
			logger.error("Found an exception: {}", e);
		}
		writer.close();
		logger.debug("Exiting {} method", methodName);
	}
}