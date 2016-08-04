import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetRequest;
import com.google.api.services.sheets.v4.model.CellData;
import com.google.api.services.sheets.v4.model.CellFormat;
import com.google.api.services.sheets.v4.model.Color;
import com.google.api.services.sheets.v4.model.CopyPasteRequest;
import com.google.api.services.sheets.v4.model.ExtendedValue;
import com.google.api.services.sheets.v4.model.GridCoordinate;
import com.google.api.services.sheets.v4.model.GridRange;
import com.google.api.services.sheets.v4.model.RepeatCellRequest;
import com.google.api.services.sheets.v4.model.Request;
import com.google.api.services.sheets.v4.model.RowData;
import com.google.api.services.sheets.v4.model.SheetProperties;
import com.google.api.services.sheets.v4.model.UpdateCellsRequest;
import com.google.api.services.sheets.v4.model.UpdateSheetPropertiesRequest;
import com.google.api.services.sheets.v4.model.ValueRange;

public class SheetsQuickstart {
	/** Application name. */
	private static final String APPLICATION_NAME = "Google Sheets API Java Quickstart";

	/** Directory to store user credentials for this application. */
	private static final java.io.File DATA_STORE_DIR 
		= new java.io.File(System.getProperty("user.home"), ".credentials/sheets.googleapis.com-java-quickstart.json");

	/** Global instance of the {@link FileDataStoreFactory}. */
	private static FileDataStoreFactory DATA_STORE_FACTORY;

	/** Global instance of the JSON factory. */
	private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();

	/** Global instance of the HTTP transport. */
	private static HttpTransport HTTP_TRANSPORT;

	/** Global instance of the scopes required by this quickstart.
	 *
	 * If modifying these scopes, delete your previously saved credentials
	 * at ~/.credentials/sheets.googleapis.com-java-quickstart.json
	 */
	//private static final List<String> SCOPES = Arrays.asList(SheetsScopes.SPREADSHEETS_READONLY);
	private static final List<String> SCOPES = Arrays.asList(SheetsScopes.SPREADSHEETS);

	static {
		try {
			HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
			DATA_STORE_FACTORY = new FileDataStoreFactory(DATA_STORE_DIR);
		} catch (Throwable t) {
			t.printStackTrace();
			System.exit(1);
		}
	}

	/**
	 * Creates an authorized Credential object.
	 * @return an authorized Credential object.
	 * @throws IOException
	 */
	public static Credential authorize() throws IOException {
		// Load client secrets.
		InputStream in = SheetsQuickstart.class.getResourceAsStream("/client_secret.json");
		GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

		// Build flow and trigger user authorization request.
		GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
				.setDataStoreFactory(DATA_STORE_FACTORY)
				.setAccessType("offline")
				.build();
		Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");
		System.out.println("Credentials saved to " + DATA_STORE_DIR.getAbsolutePath());
		return credential;
	}

	/**
	 * Build and return an authorized Sheets API client service.
	 * @return an authorized Sheets API client service
	 * @throws IOException
	 */
	public static Sheets getSheetsService() throws IOException {
		Credential credential = authorize();
		return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
				.setApplicationName(APPLICATION_NAME)
				.build();
	}

	public static void main(String[] args) throws IOException {
		editSheet2();
	}
	
	public static void getDataSheet() throws IOException{
		// Build a new authorized API client service.
		Sheets service = getSheetsService();

		// Prints the names and majors of students in a sample spreadsheet:
		// https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
		String spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
		String range = "Class Data!A2:E";
		ValueRange response = service.spreadsheets().values().get(spreadsheetId, range).execute();
		List<List<Object>> values = response.getValues();
		if (values == null || values.size() == 0) {
			System.out.println("No data found.");
		} else {
			System.out.println("Name, Major");
			for (List<Object> row : values) {
				// Print columns A and E, which correspond to indices 0 and 4.
				System.out.printf("%s, %s\n", row.get(0), row.get(4));
			}
		}
	}
	
	public static void editSheet() throws IOException{
		// Build a new authorized API client service.
		Sheets service = getSheetsService();

		// Prints the names and majors of students in a sample spreadsheet:
		// https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
		String spreadsheetId = "1AXRQ7DWUXonT4wim5wlX0CrJwUa_dTFpqi_DEnlWL6E";
		
		List<CellData> values = new ArrayList<>();
		values.add(new CellData()
				.setUserEnteredValue(new ExtendedValue().setNumberValue(Double.valueOf(1))));
		
		List<Request> requests = new ArrayList<>();
		
		requests.add(new Request().setUpdateCells(
				new UpdateCellsRequest().setStart(
						new GridCoordinate().setSheetId(0).setRowIndex(0).setColumnIndex(0))
							.setRows(Arrays.asList(new RowData().setValues(values)))
							.setFields("userEnteredValue")));

		BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest().setRequests(requests);
		service.spreadsheets().batchUpdate(spreadsheetId, batchUpdateRequest).execute();
	}
	
	public static void editSheet2() throws IOException{
		// Build a new authorized API client service.
		Sheets service = getSheetsService();

		// Prints the names and majors of students in a sample spreadsheet:
		// https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
		String spreadsheetId = "1AXRQ7DWUXonT4wim5wlX0CrJwUa_dTFpqi_DEnlWL6E";
		
		List<Request> requests = new ArrayList<>();

		// Change the name of sheet ID '0' (the default first sheet on every
		// spreadsheet)
		requests.add(new Request().setUpdateSheetProperties(new UpdateSheetPropertiesRequest()
				.setProperties(new SheetProperties().setSheetId(0).setTitle("New Sheet Name")).setFields("title")));

		// Insert the values 1, 2, 3 into the first row of the spreadsheet with
		// a different background color in each.
		List<CellData> values = new ArrayList<>();
		values.add(new CellData().setUserEnteredValue(new ExtendedValue().setNumberValue(Double.valueOf(1)))
				.setUserEnteredFormat(new CellFormat().setBackgroundColor(new Color().setRed(Float.valueOf(1)))));
		values.add(new CellData().setUserEnteredValue(new ExtendedValue().setNumberValue(Double.valueOf(2)))
				.setUserEnteredFormat(new CellFormat().setBackgroundColor(new Color().setBlue(Float.valueOf(1)))));
		values.add(new CellData().setUserEnteredValue(new ExtendedValue().setNumberValue(Double.valueOf(3)))
				.setUserEnteredFormat(new CellFormat().setBackgroundColor(new Color().setGreen(Float.valueOf(1)))));
		
		requests.add(new Request().setUpdateCells(
				new UpdateCellsRequest().setStart(new GridCoordinate().setSheetId(0).setRowIndex(0).setColumnIndex(0))
						.setRows(Arrays.asList(new RowData().setValues(values)))
						.setFields("userEnteredValue,userEnteredFormat.backgroundColor")));

		// Write "=A1+1" into A2 and fill the formula across A2:C5 (so B2 is
		// "=B1+1", C2 is "=C1+1", A3 is "=A2+1", etc..)
		requests.add(new Request().setRepeatCell(new RepeatCellRequest()
				.setRange(new GridRange().setSheetId(0).setStartRowIndex(1).setEndRowIndex(6).setStartColumnIndex(0)
						.setEndColumnIndex(3))
				.setCell(new CellData().setUserEnteredValue(new ExtendedValue().setFormulaValue("=A1 + 1")))
				.setFields("userEnteredValue")));

		// Copy the format from A1:C1 and paste it into A2:C5, so the data in
		// each column has the same background.
		requests.add(new Request().setCopyPaste(new CopyPasteRequest()
				.setSource(new GridRange().setSheetId(0).setStartRowIndex(0).setEndRowIndex(1).setStartColumnIndex(0)
						.setEndColumnIndex(3))
				.setDestination(new GridRange().setSheetId(0).setStartRowIndex(1).setEndRowIndex(6)
						.setStartColumnIndex(0).setEndColumnIndex(3))
				.setPasteType("PASTE_FORMAT")));

		BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest().setRequests(requests);
		service.spreadsheets().batchUpdate(spreadsheetId, batchUpdateRequest).execute();
	}

}